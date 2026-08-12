import os
import requests
from flask import Flask, request, abort
from linebot.v3 import WebhookHandler
from linebot.v3.exceptions import InvalidSignatureError
from linebot.v3.messaging import (
    Configuration,
    ApiClient,
    MessagingApi,
    ReplyMessageRequest,
    TextMessage
)
from linebot.v3.webhooks import MessageEvent, TextMessageContent

app = Flask(__name__)

LINE_CHANNEL_SECRET = os.getenv('LINE_CHANNEL_SECRET')
LINE_CHANNEL_ACCESS_TOKEN = os.getenv('LINE_CHANNEL_ACCESS_TOKEN')
GEMINI_API_KEY = os.getenv('GEMINI_API_KEY')

configuration = Configuration(access_token=LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

@app.route("/", methods=['GET'])
def index():
    return "B.S.S Report Bot is running!"

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers.get('X-Line-Signature')
    body = request.get_data(as_text=True)

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)

    return 'OK'

def generate_report_with_gemini(user_text, api_key):
    prompt = f"""
以下のサッカースクールのメモテキストから、保護者・生徒向けの成長レポート文章を作成してください。

入力メモ:
{user_text}

出力形式:
【成長レポート】
■ 今月の成長ポイント
(ここに150文字程度で記載)

■ 来月の目標
(ここに150文字程度で記載)
"""
    payload = {
        "contents": [
            {
                "parts": [{"text": prompt}]
            }
        ]
    }

    # APIから現在利用可能なモデル一覧を自動取得
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
    except Exception:
        pass

    # 自動取得が失敗した場合のバックアップ候補
    if not candidate_urls:
        backup_models = ["gemini-1.5-flash", "gemini-1.5-pro", "gemini-2.0-flash"]
        for bm in backup_models:
            candidate_urls.append(f"https://generativelanguage.googleapis.com/v1beta/models/{bm}:generateContent?key={api_key}")

    # 利用可能なモデルで順に呼び出しを実行
    last_error = "利用可能なGeminiモデルが見つかりませんでした。"
    for url in candidate_urls:
        try:
            res = requests.post(url, json=payload, headers={"Content-Type": "application/json"}, timeout=10)
            res_data = res.json()
            if res.status_code == 200 and 'candidates' in res_data:
                return res_data['candidates'][0]['content']['parts'][0]['text']
            else:
                last_error = res_data.get('error', {}).get('message', 'APIエラー')
        except Exception as e:
            last_error = str(e)

    return f"APIエラー: {last_error}"

@handler.add(MessageEvent, message=TextMessageContent)
def handle_message(event):
    user_text = event.message.text
    reply_text = generate_report_with_gemini(user_text, GEMINI_API_KEY)

    with ApiClient(configuration) as api_client:
        line_bot_api = MessagingApi(api_client)
        line_bot_api.reply_message(
            ReplyMessageRequest(
                reply_token=event.reply_token,
                messages=[TextMessage(text=reply_text)]
            )
        )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
