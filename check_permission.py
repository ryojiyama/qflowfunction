import msal, requests, json

# 設定読み込み（既存のlocal.settings.jsonを流用）
with open("local.settings.json", "r") as f: settings = json.load(f)["Values"]

# トークン取得
app = msal.ConfidentialClientApplication(settings["CLIENT_ID"], authority=f"https://login.microsoftonline.com/{settings['TENANT_ID']}", client_credential=settings["CLIENT_SECRET"])
token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])["access_token"]
headers = {"Authorization": f"Bearer {token}"}

# ルート直下を取得して _temp を探す
url = f"https://graph.microsoft.com/v1.0/drives/{settings['SHAREPOINT_DRIVE_ID']}/root/children"
res = requests.get(url, headers=headers).json()

for item in res.get("value", []):
    if item["name"] == "_temp":
        print(f"🎉 見つかりました！")
        print(f"   フォルダ名: {item['name']}")
        print(f"   正しいID : {item['id']}")
