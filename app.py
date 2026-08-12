import os
import json
import threading
import requests
import openpyxl
from flask import Flask, request, abort, send_file
from linebot.v3 import WebhookHandler
from linebot.v3.exceptions import InvalidSignatureError
from linebot.v3.messaging import (
    Configuration,
    ApiClient,
    MessagingApi,
    ReplyMessageRequest,
    PushMessageRequest,
    TextMessage
)
from linebot.v3.webhooks import MessageEvent, TextMessageContent

app = Flask(__name__)

LINE_CHANNEL_SECRET = os.getenv('LINE_CHANNEL_SECRET')
LINE_CHANNEL_ACCESS_TOKEN = os.getenv('LINE_CHANNEL_ACCESS_TOKEN')
GEMINI_API_KEY = os.getenv('GEMINI_API_KEY')
SERVER_BASE_URL = os.getenv('RENDER_EXTERNAL_URL', 'https://bss-report-bot.onrender.com')

configuration = Configuration(access_token=LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

TEMPLATE_PATH = "BASICサッカースクール_レポートフォーマット_A4ぴったり.xlsx"

@app.route("/", methods=['GET'])
def index():
    return "B.S.S Report Bot is running!"

@app.route("/files/<filename>", methods=['GET'])
def download_file(filename):
    file_path = os.path.join("/tmp", filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    return "File not found", 404

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers.get('X-Line-Signature')
    body = request.get_data(as_text=True)

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)

    return 'OK'

def parse_and_generate(user_text, api_key):
    prompt = f"""
以下のサッカースクールのメモテキストから、保護者・生徒向けの成長レポート情報および文章を作成し、JSON形式で出力してください。

入力メモ:
{user_text}

出力フォーマット(JSONのみ出力してください。マークダウンなどの囲みは不要です):
{{
  "name": "生徒名(テキストから抽出。無ければ空文字)",
  "grade": "学年(テキストから抽出。例: 4年。無ければ空文字)",
  "month": "月(テキストから抽出。例: 8月。無ければ空文字)",
  "course": "コース(テキストから抽出。例: 週1回。無ければ空文字)",
  "score_stop": "止めるの点数(数値または空文字)",
  "score_kick": "蹴るの点数(数値または空文字)",
  "score_carry": "運ぶの点数(数値または空文字)",
  "score_judge": "判断の点数(数値または空文字)",
  "growth_point": "今月の成長ポイント(150文字程度の丁寧な保護者向け文章)",
  "next_goal": "来月の目標(150文字程度の丁寧な保護者向け文章)"
}}
"""
    payload = {
        "contents": [{"parts": [{"text": prompt}]}]
    }

    candidate_urls = []
    try:
        list_url = f"https://generativelanguage.googleapis.com/v1beta/models?key={api_key}"
        list_res = requests.get(list_url, timeout=5)
        if list_res.status_code == 200:
            models_list = list_res.json().get("models", [])
            for m in models_list:
                if "generateContent" in m.get("supportedGenerationMethods", []):
                    m_name = m.get("name")
                    candidate_urls.append(f"https://generativelanguage.googleapis.com/v1beta/{m_name}:generateContent?key={api_key}")
    except Exception as e:
        print("Model list error:", e)

    if not candidate_urls:
        candidate_urls = [f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={api_key}"]

    for url in candidate_urls:
        try:
            res = requests.post(url, json=payload, headers={"Content-Type": "application/json"}, timeout=15)
            res_data = res.json()
            if res.status_code == 200 and 'candidates' in res_data:
                raw_text = res_data['candidates'][0]['content']['parts'][0]['text']
                raw_text = raw_text.replace("```json", "").replace("```", "").strip()
                return json.loads(raw_text)
        except Exception as e:
            print("API Post error:", e)
            continue

    return None

def create_excel_report(data):
    if not os.path.exists(TEMPLATE_PATH):
        print("Template file not found:", TEMPLATE_PATH)
        return None, None

    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    sheet = wb.active

    # セル書き込み
    if data.get("name"): sheet["C5"] = data["name"]
    if data.get("grade"): sheet["H5"] = data["grade"]
    if data.get("month"): sheet["C6"] = data["month"]
    if data.get("course"): sheet["H6"] = data["course"]

    if data.get("score_stop"): sheet["E9"] = str(data["score_stop"])
    if data.get("score_kick"): sheet["E10"] = str(data["score_kick"])
    if data.get("score_carry"): sheet["E11"] = str(data["score_carry"])
    if data.get("score_judge"): sheet["E12"] = str(data["score_judge"])

    if data.get("growth_point"): sheet["B16"] = data["growth_point"]
    if data.get("next_goal"): sheet["B25"] = data["next_goal"]

    file_name = f"成長レポート_{data.get('name', '生徒')}_{data.get('month', '')}.xlsx"
    file_path = os.path.join("/tmp", file_name)
    wb.save(file_path)
    return file_name, file_path

def async_process_and_reply(reply_token, user_text):
    parsed_data = parse_and_generate(user_text, GEMINI_API_KEY)

    if not parsed_data:
        reply_messages = [TextMessage(text="申し訳ありません。レポートデータの生成に失敗しました。")]
    else:
        file_name, file_path = create_excel_report(parsed_data)
        
        if file_name and file_path:
            file_url = f"{SERVER_BASE_URL}/files/{file_name}"
            summary_text = f"【{parsed_data.get('name', '生徒')}さんの成長レポートを作成しました】\n\n" \
                           f"■今月の成長ポイント\n{parsed_data.get('growth_point', '')}\n\n" \
                           f"■来月の目標\n{parsed_data.get('next_goal', '')}\n\n" \
                           f"📥 完成したExcelファイルのダウンロード:\n{file_url}"
        else:
            summary_text = f"【{parsed_data.get('name', '生徒')}さんの成長レポートを作成しました】\n\n" \
                           f"■今月の成長ポイント\n{parsed_data.get('growth_point', '')}\n\n" \
                           f"■来月の目標\n{parsed_data.get('next_goal', '')}\n\n" \
                           f"※Excelテンプレートが見つからなかったため、文章のみ出力しました。"

        reply_messages = [TextMessage(text=summary_text)]

    try:
        with ApiClient(configuration) as api_client:
            line_bot_api = MessagingApi(api_client)
            line_bot_api.reply_message(
                ReplyMessageRequest(
                    reply_token=reply_token,
                    messages=reply_messages
                )
            )
        print("Successfully replied to LINE!")
    except Exception as e:
        print("Failed to reply to LINE:", e)

@handler.add(MessageEvent, message=TextMessageContent)
def handle_message(event):
    user_text = event.message.text
    reply_token = event.reply_token

    # バックグラウンドスレッドで重い処理（AI・Excel作成）を実行
    thread = threading.Thread(target=async_process_and_reply, args=(reply_token, user_text))
    thread.start()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
