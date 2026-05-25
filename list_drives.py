import msal, requests, json

# 既存の設定を読み込む
with open("local.settings.json", "r") as f:
    settings = json.load(f)["Values"]

# 認証トークンの取得
app = msal.ConfidentialClientApplication(
    settings["CLIENT_ID"],
    authority=f"https://login.microsoftonline.com/{settings['TENANT_ID']}",
    client_credential=settings["CLIENT_SECRET"]
)
token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])["access_token"]

# サイト内の全ライブラリ（ドライブ）一覧を取得
site_id = settings['SHAREPOINT_SITE_ID']
url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
res = requests.get(url, headers={"Authorization": f"Bearer {token}"}).json()

print("--- サイト内のライブラリ一覧 ---")
for drive in res.get("value", []):
    print(f"名前: {drive['name']} | ID: {drive['id']}")
