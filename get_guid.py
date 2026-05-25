import os
import json
import msal
import requests

# local.settings.json から資格情報を自動読み込み
with open("local.settings.json", "r", encoding="utf-8") as f:
    settings = json.load(f)
    env = settings.get("Values", {})

TENANT_ID = env.get("TENANT_ID")
CLIENT_ID = env.get("CLIENT_ID")
CLIENT_SECRET = env.get("CLIENT_SECRET")

# トークン取得
authority = f"https://login.microsoftonline.com/{TENANT_ID}"
app = msal.ConfidentialClientApplication(CLIENT_ID, authority=authority, client_credential=CLIENT_SECRET)
token_res = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
token = token_res.get("access_token")
headers = {"Authorization": f"Bearer {token}"}

print("1. Site IDを取得中...")
# 検索ではなく、URLのパスから直接サイトを特定します
site_url = "https://graph.microsoft.com/v1.0/sites/tshldgs.sharepoint.com:/sites/QualityAssurance"
site_res = requests.get(site_url, headers=headers)

if site_res.status_code != 200:
    print(f"エラー: サイトの取得に失敗しました。詳細: {site_res.text}")
    exit()

site = site_res.json()
site_id = site['id']
print(f"✅ サイト名: {site.get('displayName', 'Quality Assurance')}")
print(f"✅ SHAREPOINT_SITE_ID: {site_id}\n")

print("2. Drive ID (ドキュメントライブラリ) を取得中...")
drive_res = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive", headers=headers)

if drive_res.status_code != 200:
    print(f"エラー: ドライブの取得に失敗しました。詳細: {drive_res.text}")
    exit()

drive_id = drive_res.json().get("id")
print(f"✅ SHAREPOINT_DRIVE_ID: {drive_id}\n")

print("★ 上記の2つのIDを local.settings.json にコピーしてください！")
