import requests
import json

# ローカルで起動したAzure FunctionsのURL
url = "http://localhost:7071/api/StampPDF"

# SOW v4 に基づくリクエストボディ
payload = {
    "distribution_id": "999",  # テスト用のダミーID（リストアイテムのID）
    "temp_file_url": "https://tshldgs.sharepoint.com/sites/QualityAssurance/Shared%20Documents/_temp/テスト用ファイル.docx" # ★先ほどアップロードしたWord/ExcelのURLに変更してください
}

print("Azure Functions に指示を送信中...")
response = requests.post(url, json=payload)

print(f"HTTP Status: {response.status_code}")
try:
    print("レスポンス内容:", json.dumps(response.json(), indent=2, ensure_ascii=False))
except:
    print(response.text)
