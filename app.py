import os
import json
import base64
import re
import traceback
import io
from flask import Flask, request, abort
from PIL import Image

# LINE SDK v3
from linebot.v3 import WebhookHandler
from linebot.v3.exceptions import InvalidSignatureError
from linebot.v3.messaging import (
    Configuration,
    ApiClient,
    MessagingApi,
    MessagingApiBlob,
    ReplyMessageRequest,
    PushMessageRequest,
    TextMessage,
)
from linebot.v3.webhooks import (
    MessageEvent,
    TextMessageContent,
    ImageMessageContent,
)

import anthropic
import pandas as pd

app = Flask(__name__)

# ===== 認証情報 =====
LINE_CHANNEL_SECRET = os.environ.get('LINE_CHANNEL_SECRET', 'e89ce78ff68ccbed4256eaf8a6b1e956')
LINE_CHANNEL_ACCESS_TOKEN = os.environ.get(
    'LINE_CHANNEL_ACCESS_TOKEN',
    'oDIr4S9ApUh7pMoBr8ysJBhWLD3WWmFAWguITuDIoCY9ald3pTL4Kow3nGEyo7YZ/QXlC2GGsbFXaFKUhH7JhHCE+0yd63V4d0n7PE1M2KtQ3E0cfgCTwib6V9R1dzpZpAvyuCN8nh414PZpOcCi6wdB04t89/1O/w1cDnyilFU='
)
ANTHROPIC_API_KEY = os.environ.get('ANTHROPIC_API_KEY', '')

# LINE SDK 設定
configuration = Configuration(access_token=LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)
anthropic_client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)


# ===== マスターデータ読み込み =====
def load_master_data():
    """Excelマスター価格表を読み込む"""
    excel_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '競合価格比較.xlsx')
    df = pd.read_excel(excel_path)
    # 列名の空白を除去
    df.columns = df.columns.str.strip()
    # 価格列を数値に変換
    df['自社本体価格(税抜)'] = pd.to_numeric(df['自社本体価格(税抜)'], errors='coerce')
    # 有効なデータだけ残す
    df = df.dropna(subset=['商品名', '自社本体価格(税抜)'])
    print(f"マスターデータ読み込み完了: {len(df)}件")
    return df


# 起動時に読み込み
master_df = load_master_data()


# ===== 商品名マッチング =====
def find_matching_item(flyer_name: str, df: pd.DataFrame):
    """チラシの商品名とマスターデータを照合する"""
    flyer_name_clean = flyer_name.strip()

    # 完全一致
    exact = df[df['商品名'] == flyer_name_clean]
    if len(exact) > 0:
        return exact.iloc[0]

    # 部分一致（チラシ名がマスター名に含まれる、またはその逆）
    for _, row in df.iterrows():
        master_name = str(row['商品名'])
        if flyer_name_clean in master_name or master_name in flyer_name_clean:
            return row

    # キーワード一致（2文字以上の単語で照合）
    best_match = None
    best_score = 0
    words = [w for w in re.split(r'[\s　・/]+', flyer_name_clean) if len(w) >= 2]

    for _, row in df.iterrows():
        master_name = str(row['商品名'])
        score = sum(1 for w in words if w in master_name)
        if score > best_score and score >= 2:
            best_score = score
            best_match = row

    return best_match


# ===== メッセージ送信ヘルパー =====
def reply_message(reply_token: str, message_text: str):
    with ApiClient(configuration) as api_client:
        line_bot_api = MessagingApi(api_client)
        line_bot_api.reply_message(
            ReplyMessageRequest(
                reply_token=reply_token,
                messages=[TextMessage(text=message_text)]
            )
        )


def push_message(user_id: str, message_text: str):
    try:
        with ApiClient(configuration) as api_client:
            line_bot_api = MessagingApi(api_client)
            line_bot_api.push_message(
                PushMessageRequest(
                    to=user_id,
                    messages=[TextMessage(text=message_text)]
                )
            )
    except Exception as e:
        print(f"Push message error: {e}")


# ===== LINE Webhook エンドポイント =====
@app.route('/webhook', methods=['POST'])
def webhook():
    signature = request.headers.get('X-Line-Signature', '')
    body = request.get_data(as_text=True)

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        print("Invalid signature")
        abort(400)

    return 'OK'


# ===== 画像メッセージ処理 =====
@handler.add(MessageEvent, message=ImageMessageContent)
def handle_image_message(event):
    """チラシ画像を受け取り、負け商品を返信する"""

    user_id = event.source.user_id
    message_id = event.message.id

    # まず「処理中」と返信（reply tokenを使用）
    reply_message(event.reply_token, "📸 チラシを解析中です...\n少しお待ちください（10〜20秒）")

    try:
        # 画像データ取得
        with ApiClient(configuration) as api_client:
            blob_api = MessagingApiBlob(api_client)
            image_data = blob_api.get_message_content(message_id)

        # 画像を圧縮（大きすぎるとAPIエラーになるため）
        try:
            img = Image.open(io.BytesIO(image_data))
            # 最大1600px以内にリサイズ
            max_size = 1600
            if img.width > max_size or img.height > max_size:
                img.thumbnail((max_size, max_size), Image.LANCZOS)
            # JPEG形式で圧縮
            output = io.BytesIO()
            if img.mode in ('RGBA', 'P'):
                img = img.convert('RGB')
            img.save(output, format='JPEG', quality=85)
            image_data = output.getvalue()
            print(f"画像圧縮完了: {len(image_data)} bytes")
        except Exception as img_err:
            print(f"画像圧縮スキップ: {img_err}")

        # base64エンコード
        image_base64 = base64.standard_b64encode(image_data).decode('utf-8')

        # 画像タイプ判定（圧縮後はJPEG）
        media_type = 'image/jpeg'

        # Claude Vision APIで商品名と価格を抽出
        claude_response = anthropic_client.messages.create(
            model='claude-haiku-4-5-20251001',
            max_tokens=3000,
            messages=[
                {
                    'role': 'user',
                    'content': [
                        {
                            'type': 'image',
                            'source': {
                                'type': 'base64',
                                'media_type': media_type,
                                'data': image_base64,
                            }
                        },
                        {
                            'type': 'text',
                            'text': """このスーパーのチラシ画像から、商品名と価格をすべて抽出してください。

以下のJSON形式のみで返答してください（説明文は不要）：
{"items": [{"name": "商品名", "price": 価格の数値}, ...]}

ルール：
- 価格は税抜き本体価格（円）を数値で記載
- 税込み価格しか読み取れない場合は1.1で割って税抜きに変換（小数点以下切り捨て）
- 「本体」「税抜」と書いてある価格を優先
- 商品名はチラシに書いてある通りに記載
- 価格が読み取れない商品は除外"""
                        }
                    ]
                }
            ]
        )

        result_text = claude_response.content[0].text.strip()

        # JSONを抽出
        json_match = re.search(r'\{[\s\S]*\}', result_text)
        if not json_match:
            push_message(user_id, "❌ チラシから商品情報を読み取れませんでした。\n明るい場所で撮り直してください。")
            return

        items_data = json.loads(json_match.group())
        flyer_items = items_data.get('items', [])

        if not flyer_items:
            push_message(user_id, "❌ チラシから商品が見つかりませんでした。\nもっと近づいて撮影してください。")
            return

        # マスターデータと照合
        losing_items = []
        matched_count = 0

        for flyer_item in flyer_items:
            flyer_name = str(flyer_item.get('name', '')).strip()
            flyer_price_raw = flyer_item.get('price', 0)

            if not flyer_name or not flyer_price_raw:
                continue

            try:
                flyer_price = float(flyer_price_raw)
            except (ValueError, TypeError):
                continue

            matched_row = find_matching_item(flyer_name, master_df)

            if matched_row is not None:
                matched_count += 1
                our_price = float(matched_row['自社本体価格(税抜)'])

                if our_price > flyer_price:
                    diff = our_price - flyer_price
                    losing_items.append({
                        'master_name': str(matched_row['商品名']),
                        'our_price': int(our_price),
                        'competitor_price': int(flyer_price),
                        'diff': int(diff)
                    })

        # 返信メッセージ作成
        total_flyer = len(flyer_items)

        if losing_items:
            losing_items.sort(key=lambda x: x['diff'], reverse=True)

            message = f"🔴 負け商品: {len(losing_items)}件！\n"
            message += f"（チラシ{total_flyer}品中 {matched_count}品照合）\n"
            message += "━━━━━━━━━━━━━━\n"

            for item in losing_items[:15]:
                message += f"\n❌ {item['master_name']}\n"
                message += f"   当店¥{item['our_price']} → 競合¥{item['competitor_price']} (-¥{item['diff']})\n"

            if len(losing_items) > 15:
                message += f"\n他{len(losing_items) - 15}件あります"
        else:
            if matched_count == 0:
                message = f"⚠️ マスターデータと一致する商品が見つかりませんでした。\n（チラシ{total_flyer}品を確認）"
            else:
                message = f"✅ 負け商品なし！\n（チラシ{total_flyer}品中 {matched_count}品照合）"

        push_message(user_id, message)

    except json.JSONDecodeError as e:
        print(f"JSON parse error: {e}")
        push_message(user_id, "❌ AIの返答を解析できませんでした。\nもう一度送信してください。")
    except Exception as e:
        error_detail = traceback.format_exc()
        print(f"Error: {error_detail}")
        # エラーの種類をユーザーに通知
        error_short = str(e)[:150]
        push_message(user_id, f"❌ エラーが発生しました。\n詳細: {error_short}")


# ===== テキストメッセージ処理 =====
@handler.add(MessageEvent, message=TextMessageContent)
def handle_text_message(event):
    text = event.message.text.strip()

    if text in ['ヘルプ', 'help', '使い方', '？', '?']:
        reply_text = (
            "📊 競合価格比較ボット\n\n"
            "【使い方】\n"
            "チラシの写真を送ると\n当店より安い商品を自動検出します\n\n"
            "【手順】\n"
            "① チラシをスマホで撮影\n"
            "② この画面に画像を送信\n"
            "③ 10〜20秒で結果が届く\n\n"
            f"登録商品数: {len(master_df)}件"
        )
    else:
        reply_text = "チラシの写真を送ってください📸\n「ヘルプ」で使い方を確認できます。"

    reply_message(event.reply_token, reply_text)


# ===== ヘルスチェック =====
@app.route('/', methods=['GET'])
def health_check():
    return f'競合価格比較ボット 稼働中 (マスター{len(master_df)}件)'


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
