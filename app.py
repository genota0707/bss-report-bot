import os
import io
import re
from flask import Flask, request, abort, send_file
from linebot.v3 import WebhookHandler
from linebot.v3.exceptions import InvalidSignatureError
from linebot.v3.messaging import (
    Configuration,
    ApiClient,
    MessagingApi,
    MessagingApiBlob,
    ReplyMessageRequest,
    TextMessage,
    PushMessageRequest
)
from linebot.v3.webhooks import MessageEvent, TextMessageContent
import google.generativeai as genai
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = Flask(__name__)

# 環境変数から設定を取得
LINE_CHANNEL_SECRET = os.getenv('LINE_CHANNEL_SECRET')
LINE_CHANNEL_ACCESS_TOKEN = os.getenv('LINE_CHANNEL_ACCESS_TOKEN')
GEMINI_API_KEY = os.getenv('GEMINI_API_KEY')

# API初期化
configuration = Configuration(access_token=LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)
genai.configure(api_key=GEMINI_API_KEY)

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

def create_excel_report(data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "レポート"

    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.25
    ws.page_margins.bottom = 0.25

    font_title = Font(name="メイリオ", size=22, bold=True, color="FFFFFF")
    font_subtitle = Font(name="メイリオ", size=12, italic=True, color="E0E8F0")
    font_label_bold = Font(name="メイリオ", size=13, bold=True, color="1B2A4A")
    font_header_sec = Font(name="メイリオ", size=13, bold=True, color="FFFFFF")
    font_body = Font(name="メイリオ", size=11, color="000000")
    font_small = Font(name="メイリオ", size=10, color="444444")

    fill_header = PatternFill(start_color="1B2A4A", end_color="1B2A4A", fill_type="solid")
    fill_sec_header = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    fill_light_gray = PatternFill(start_color="F4F6F9", end_color="F4F6F9", fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    thin_border_side = Side(border_style="thin", color="CCCCCC")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    col_widths = {'A': 3, 'B': 12, 'C': 12, 'D': 10, 'E': 10, 'F': 14, 'G': 10, 'H': 12, 'I': 14, 'J': 10, 'K': 4}
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

    ws.row_dimensions[2].height = 36
    ws.row_dimensions[3].height = 24
    ws.merge_cells('B2:J2')
    ws['B2'] = "B.S.S"
    ws['B2'].font = font_title
    ws['B2'].alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws['B2'].fill = fill_header

    ws.merge_cells('B3:J3')
    ws['B3'] = "BASIC soccer school"
    ws['B3'].font = font_subtitle
    ws['B3'].alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws['B3'].fill = fill_header

    ws.row_dimensions[5].height = 32
    ws.row_dimensions[6].height = 32
    for r in range(5, 7):
        for c in range(2, 11):
            cell = ws.cell(r, c)
            cell.fill = fill_light_gray
            cell.border = thin_border

    ws['B5'] = f"生徒名： {data.get('name', '')}"
    ws['B5'].font = font_label_bold
    ws['B5'].alignment = Alignment(horizontal='left', vertical='center', indent=1)

    ws['G5'] = f"学年： {data.get('grade', '')}"
    ws['G5'].font = font_label_bold
    ws['G5'].alignment = Alignment(horizontal='left', vertical='center')

    ws['B6'] = f"月： {data.get('month', '')}"
    ws['B6'].font = font_label_bold
    ws['B6'].alignment = Alignment(horizontal='left', vertical='center', indent=1)

    ws['G6'] = f"コース： {data.get('course', '')}"
    ws['G6'].font = font_label_bold
    ws['G6'].alignment = Alignment(horizontal='left', vertical='center')

    ws.row_dimensions[8].height = 28
    ws.merge_cells('B8:J8')
    ws['B8'] = "■ 評価項目"
    ws['B8'].font = font_header_sec
    ws['B8'].fill = fill_sec_header
    ws['B8'].alignment = Alignment(horizontal='left', vertical='center', indent=1)

    scores = data.get('scores', {'止める': '', '蹴る': '', '運ぶ': '', '判断': ''})
    items = [("止める", 9), ("蹴る", 10), ("運ぶ", 11), ("判断", 12)]
    
    for item, idx in items:
        ws.row_dimensions[idx].height = 34
        ws.cell(row=idx, column=2, value=item).font = font_label_bold
        ws.cell(row=idx, column=2).alignment = Alignment(horizontal='center', vertical='center')
        
        val = scores.get(item, '')
        ws.merge_cells(start_row=idx, start_column=3, end_row=idx, end_column=4)
        ws.cell(row=idx, column=3, value=f"{val} / 10" if val else "/ 10").font = font_label_bold
        ws.cell(row=idx, column=3).alignment = Alignment(horizontal='left', vertical='center')
        
        for c in range(2, 11):
            cell = ws.cell(idx, c)
            cell.border = thin_border
            if idx % 2 == 1:
                cell.fill = fill_light_gray

    ws.row_dimensions[13].height = 24
    ws.merge_cells('B13:J13')
    ws['B13'] = "項目： / 10   2 ~ 4 → 🌱   5 ~ 7 → 🌿   8 ~ 10 → 🌳"
    ws['B13'].font = font_small
    ws['B13'].alignment = Alignment(horizontal='left', vertical='center', indent=1)

    ws.row_dimensions[15].height = 28
    ws.merge_cells('B15:J15')
    ws['B15'] = "今月の成長ポイント："
    ws['B15'].font = font_header_sec
    ws['B15'].fill = fill_sec_header
    ws['B15'].alignment = Alignment(horizontal='left', vertical='center', indent=1)

    ws.merge_cells('B16:J22')
    for r in range(16, 23):
        ws.row_dimensions[r].height = 24
        for c in range(2, 11):
            cell = ws.cell(r, c)
            cell.border = thin_border
            cell.fill = fill_white
    ws['B16'] = data.get('growth', '')
    ws['B16'].font = font_body
    ws['B16'].alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

    ws.row_dimensions[24].height = 28
    ws.merge_cells('B24:J24')
    ws['B24'] = "来月の目標："
    ws['B24'].font = font_header_sec
    ws['B24'].fill = fill_sec_header
    ws['B24'].alignment = Alignment(horizontal='left', vertical='center', indent=1)

    ws.merge_cells('B25:J31')
    for r in range(25, 32):
        ws.row_dimensions[r].height = 24
        for c in range(2, 11):
            cell = ws.cell(r, c)
            cell.border = thin_border
            cell.fill = fill_white
    ws['B25'] = data.get('goal', '')
    ws['B25'].font = font_body
    ws['B25'].alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

@handler.add(MessageEvent, message=TextMessageContent)
def handle_message(event):
    user_text = event.message.text

    # Geminiでメモを解析＆文章作成
    model = genai.GenerativeModel('gemini-1.5-flash')
    prompt = f"""
以下のサッカースクールのメモテキストから、保護者・生徒向けの成長レポート用の情報を抽出・生成してください。

入力メモ:
{user_text}

出力フォーマット（必ず以下の形式を守って出力してください）:
生徒名: [名前]
学年: [学年]
月: [対象月]
コース: [コース]
止める: [数値]
蹴る: [数値]
運ぶ: [数値]
判断: [数値]
---成長ポイント---
[保護者・生徒に向けた温かい今月の成長ポイント文章（150〜200文字程度）]
---来月の目標---
[保護者・生徒に向けた温かい来月の目標文章（150〜200文字程度）]
"""

    response = model.generate_content(prompt)
    res_text = response.text

    # 情報パース
    data = {'scores': {}}
    for line in res_text.splitlines():
        if line.startswith('生徒名:'): data['name'] = line.split(':', 1)[1].strip()
        elif line.startswith('学年:'): data['grade'] = line.split(':', 1)[1].strip()
        elif line.startswith('月:'): data['month'] = line.split(':', 1)[1].strip()
        elif line.startswith('コース:'): data['course'] = line.split(':', 1)[1].strip()
        elif line.startswith('止める:'): data['scores']['止める'] = line.split(':', 1)[1].strip()
        elif line.startswith('蹴る:'): data['scores']['蹴る'] = line.split(':', 1)[1].strip()
        elif line.startswith('運ぶ:'): data['scores']['運ぶ'] = line.split(':', 1)[1].strip()
        elif line.startswith('判断:'): data['scores']['判断'] = line.split(':', 1)[1].strip()

    if '---成長ポイント---' in res_text and '---来月の目標---' in res_text:
        parts = res_text.split('---成長ポイント---')[1].split('---来月の目標---')
        data['growth'] = parts[0].strip()
        data['goal'] = parts[1].strip()

    # LINEに応答メッセージを返信
    reply_text = f"【{data.get('name', '生徒')}くんのレポート作成完了】\n\n■ 今月の成長ポイント:\n{data.get('growth', '')}\n\n■ 来月の目標:\n{data.get('goal', '')}\n\n※Excelレポートファイルを生成しました！"

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
