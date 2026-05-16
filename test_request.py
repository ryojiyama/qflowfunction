import requests
import base64
import json
import os

# ローカルFunctionsのURL
# url = "http://localhost:7071/api/qflowTrigger"
# デプロイ後のAzure FunctionsのURL
url = os.environ.get("AZURE_FUNCTION_URL")

# 1. PDFファイルを読み込んでBase64に変換
with open("test_Vertical.pdf", "rb") as f:
    pdf_base64 = base64.b64encode(f.read()).decode('utf-8')

# 2. テスト用のメタデータを作成
payload = {
    "pdf_content": pdf_base64,
    "metadata": {
        "dept": "QA-DIV",
        "product": "PAPR-S",
        "category": "SPEC",
        "year": "2026",
        "seq": "001",
        "rev_date": "2026/05/14",
        "confidential_level": "Internal Only",
        "qr_url": "https://example.com/test-id-123"
    }
}

# 3. POSTリクエストを送信
print("Sending request to local Azure Functions...")
response = requests.post(url, json=payload)

if response.status_code == 200:
    # 4. 返ってきた加工済みのPDFを保存
    with open("output_stamped.pdf", "wb") as f:
        f.write(response.content)
    print("成功！ 'output_stamped.pdf' が作成されました。")
else:
    print(f"失敗: {response.status_code}")
    print(response.text)
