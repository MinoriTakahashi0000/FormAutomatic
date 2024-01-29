import re
import datetime
import json
import os.path
import webbrowser
from flask import Flask, render_template, request, redirect, url_for, session
from flask_debugtoolbar import DebugToolbarExtension
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

app = Flask(__name__)
app.secret_key = "skldhjnvjlajhrkasmvkl34r89jl"
app.debug = True
toolbar = DebugToolbarExtension(app)


# urlからidを抽出する関数
def extract_id(url):
    pattern = r"/spreadsheets/d/([a-zA-Z0-9-_]+)"
    match = re.search(pattern, url)

    if match:
        return match.group(1)
    else:
        return None


# シートデータを取得する関数
def get_sheets_data(sheet_id):
    # 対象となるスプレッドシートのIDと読み取り範囲
    SHEET_ID = sheet_id
    SHEET_NAME = "フォームの回答 1"

    try:
        credentials_json = session['credentials']
        credentials = credentials.from_authorized_user_info(json.loads(credentials_json))
    
        service_sheets = build("sheets", "v4", credentials = credentials)

        spreadsheet = (
            service_sheets.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
        )
        title = spreadsheet["properties"]["title"]

        # Call the Sheets API
        result = (
            service_sheets.spreadsheets()
            .values()
            .get(spreadsheetId=SHEET_ID, range=SHEET_NAME)
            .execute()
        )
        values = result.get("values", [])

        return values, title

    except HttpError as err:
        print(err)
        return None

# スコープ
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/documents",
]

# Google Cloud Consoleで設定したリダイレクトURI
REDIRECT_URI = 'https://formautomatic.onrender.com/oauth2callback'

@app.route('/auth')
def auth():
    # 環境変数から認証情報を読み込む
    credentials_info = json.loads(os.getenv('GOOGLE_CREDENTIALS')) 
    # Flowオブジェクトを初期化
    flow = Flow.from_client_config(
        client_config=credentials_info,
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI)
    # 認証URLを生成
    authorization_url, state = flow.authorization_url(
        access_type='offline',
        include_granted_scopes='true')
    # 状態をセッションに保存
    session['state'] = state
    # ユーザーを認証URLにリダイレクト
    return redirect(authorization_url)


@app.route('/oauth2callback')
def oauth2callback():
    # 環境変数から認証情報を読み込む
    credentials_info = json.loads(os.getenv('GOOGLE_CREDENTIALS'))
    # Flowオブジェクトを再初期化
    flow = Flow.from_client_config(
        client_config=credentials_info,
        scopes=SCOPES,
        state=session['state'],
        redirect_uri=REDIRECT_URI)
    # リクエストから認証コードを取得し、アクセストークンに交換
    flow.fetch_token(authorization_response=request.url)  
    # flow.credentialsにアクセストークンとリフレッシュトークンが含まれる
    credentials = flow.credentials
    credentials_json = credentials.to_json()
    session['credentials'] = credentials_json
    # 認証が完了した後のリダイレクト先にリダイレクト
    return redirect(url_for('index'))


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process():
    url = request.form["url_input"]
    id = extract_id(url)
    sheets_data, sheets_title = get_sheets_data(id)
    # 質問取り出す
    keys = json.dumps(sheets_data[0])
    sheets_data = json.dumps(sheets_data)

    return redirect(
        url_for(
            "results", sheets_data=sheets_data, sheets_title=sheets_title, keys=keys
        )
    )


@app.route("/results")
def results():
    sheets_data = request.args.get("sheets_data", "Unknown")
    sheets_title = request.args.get("sheets_title", "Unknown")
    keys = request.args.get("keys", "Unknown")
    keys = json.loads(keys)
    return render_template(
        "results.html", sheets_title=sheets_title, keys=keys, sheets_data=sheets_data
    )


@app.route("/create_document", methods=["GET", "POST"])
def write_to_google_doc():
    # POSTリクエストからJSONデータを取得
    request_data = request.get_json()
    sheets_data = request_data.get("requestData", {}).get("sheets_data")
    sheets_data = json.loads(sheets_data)
    title = request_data.get("requestData", {}).get("title")
    selected_keys = request_data.get("requestData", {}).get("selectedKeys")

    new_list = []
    for i in range(len(sheets_data[0])):
        question_data = [sheets_data[0][i]]
        for j in range(1, len(sheets_data)):
            if i < len(sheets_data[j]):
                question_data.append(sheets_data[j][i])
            else:
                question_data.append("")
        new_list.append(question_data)

    count_answer = sum(isinstance(item, list) for item in sheets_data) - 1

    selected_list = [new_list[i] for i in range(len(new_list)) if selected_keys[i]]
    selected_list = [[item for item in row if item != ""] for row in selected_list]

    try:
        credentials_json = session['credentials']
        credentials = credentials.from_authorized_user_info(json.loads(credentials_json))
        service = build("docs", "v1", credentials=credentials)
        body = {"title": title}
        doc = service.documents().create(body=body).execute()
        document_id = doc.get("documentId")
        document_url = f"https://docs.google.com/document/d/{document_id}/edit"

        requests = []

        # ドキュメントを逆順に構築
        for line_of_text in reversed(selected_list):
            bolded_text = f"＜{line_of_text[0]}＞\n"
            line = "_________________________________________________________________________\n\n"
            result_str = ""
            filtered_list = [x for x in line_of_text if x is not None]
            for i in range(len(filtered_list[1:])):
                result_str += f"・{filtered_list[1:][i]}\n"

            # 背景色を設定
            sub_text_style = {
                "bold": False,
            }

            requests.extend(
                [
                    {
                        "insertText": {
                            "text": bolded_text,
                            "location": {
                                "index": 1,
                            },
                        },
                    },
                    {
                        "updateTextStyle": {
                            "textStyle": {
                                "bold": True,
                            },
                            "fields": "bold",
                            "range": {
                                "startIndex": 1,
                                "endIndex": len(bolded_text),
                            },
                        },
                    },
                    {
                        "insertText": {
                            "text": result_str + "\n",
                            "location": {
                                "index": len(bolded_text) + 1,
                            },
                        },
                    },
                    {
                        "updateTextStyle": {
                            "textStyle": sub_text_style,
                            "fields": "bold,italic,backgroundColor",
                            "range": {
                                "startIndex": len(bolded_text) + 1,
                                "endIndex": len(bolded_text) + len(result_str),
                            },
                        },
                    },
                ]
            )

        requests.extend(
            [
                {
                    "insertText": {
                        "text": line,
                        "location": {
                            "index": 1,
                        },
                    },
                },
            ]
        )

        sub_header_text = (
            "回答者数："
            + str(count_answer)
            + "\n"
            + "アンケート作成日："
            + str(datetime.datetime.now().date())
            + "\n"
        )
        requests.extend(
            [
                {
                    "insertText": {
                        "text": sub_header_text,
                        "location": {
                            "index": 1,
                        },
                    },
                },
                {
                    "updateTextStyle": {
                        "textStyle": {
                            "bold": False,
                            "italic": False,
                            "fontSize": {
                                "magnitude": 11,
                                "unit": "PT",
                            },
                        },
                        "fields": "bold,italic,fontSize",
                        "range": {
                            "startIndex": 1,
                            "endIndex": len(sub_header_text),
                        },
                    },
                },
            ]
        )

        requests.extend(
            [
                {
                    "insertText": {
                        "text": line,
                        "location": {
                            "index": 1,
                        },
                    },
                },
            ]
        )

        doc_header_text = title + "\n\n"
        requests.extend(
            [
                {
                    "insertText": {
                        "text": doc_header_text,
                        "location": {
                            "index": 1,
                        },
                    },
                },
                {
                    "updateTextStyle": {
                        "textStyle": {
                            "bold": False,
                            "italic": False,
                            "fontSize": {
                                "magnitude": 17,
                                "unit": "PT",
                            },
                        },
                        "fields": "bold,italic,fontSize",
                        "range": {
                            "startIndex": 1,
                            "endIndex": len(doc_header_text),
                        },
                    },
                },
                {
                    "updateParagraphStyle": {
                        "range": {
                            "startIndex": 1,
                            "endIndex": len(doc_header_text),
                        },
                        "paragraphStyle": {
                            "alignment": "CENTER",
                        },
                        "fields": "alignment",
                    },
                },
            ]
        )

        service.documents().batchUpdate(
            documentId=document_id, body={"requests": requests}
        ).execute()
        webbrowser.open(document_url, new=2)
        return redirect(url_for("end"))
    except HttpError as err:
        print(err)


@app.route("/end")
def end():
    return render_template("end.html")


if __name__ == "__main__":
    app.run(port=8000, debug=True)
