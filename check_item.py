import os
import json
import requests
import msal
import urllib.parse

# 確認したいPDFのファイル名を入力してください
TARGET_FILE_NAME = "テスト用ファイル.pdf"

def main():
    # local.settings.json から環境変数を読み込み
    with open("local.settings.json", "r", encoding="utf-8") as f:
        data = json.load(f)
        for k, v in data.get("Values", {}).items():
            os.environ[k] = str(v)

    # トークン取得
    tenant_id = os.environ.get("TENANT_ID")
    client_id = os.environ.get("CLIENT_ID")
    client_secret = os.environ.get("CLIENT_SECRET")
    app = msal.ConfidentialClientApplication(client_id, authority=f"https://login.microsoftonline.com/{tenant_id}", client_credential=client_secret)
    token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"]).get("access_token")
    headers = {"Authorization": f"Bearer {token}"}

    qpub_id = os.environ.get("QPUBLISHED_DRIVE_ID")
    encoded_name = urllib.parse.quote(TARGET_FILE_NAME)

    print(f"🔍 Q-Published内の '{TARGET_FILE_NAME}' のメタデータを調査中...\n")

    # ファイルのリストアイテム属性（メタデータ）を取得
    url = f"https://graph.microsoft.com/v1.0/drives/{qpub_id}/root:/{encoded_name}:/listitem?expand=fields"
    res = requests.get(url, headers=headers)

    if res.status_code == 200:
        fields = res.json().get("fields", {})
        # 見やすくフォーマットして出力
        print(json.dumps(fields, indent=2, ensure_ascii=False))
    else:
        print(f"❌ 取得失敗: {res.status_code}\n{res.text}")

if __name__ == "__main__":
    main()
