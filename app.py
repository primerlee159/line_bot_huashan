from __future__ import unicode_literals
from flask import Flask, request, abort

from linebot import (
    LineBotApi, WebhookHandler
)
from linebot.exceptions import (
    InvalidSignatureError
)
from linebot.models import *

app = Flask(__name__)

# Channel Access Token
line_bot_api = LineBotApi('AXdHTxXqf0Ow2BWMcBzgZmeKZuGH/Br2EwKP/7Lm+ZjIgFR34rTETLz1CrX6joBKHStI7USsP1WrkOHQ5qj317OwVoD6+RaAgtQU5TDeLOHQ+f5U5m+LO7XNZb6SJuTP42En/LoX0ac0AgGhNWyo1wdB04t89/1O/w1cDnyilFU=')
# Channel Secret
handler = WebhookHandler('9f4d78a2b22fb48a7ba59bcee7fa74cd')

# 監聽所有來自 /callback 的 Post Request
@app.route("/callback", methods=['POST'])
def callback():
    # get X-Line-Signature header value
    signature = request.headers['X-Line-Signature']
    # get request body as text
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)
    # handle webhook body
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(200)
    return 'OK'

# 處理訊息
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    # message = TextSendMessage(text=event.message.text)
    message = "收到"
    line_bot_api.reply_message(event.reply_token, message)

import os
if __name__ == "__main__":
    # port = int(os.environ.get('PORT', 80))
    # app.run(host='0.0.0.0', port=port)
    app.run()
    