import msal, requests, json

# 1. 設定の読み込み
with open("local.settings.json", "r") as f:
    settings = json.load(f)["Values"]

# 2. トークンの取得
app = msal.ConfidentialClientApplication(
    settings["CLIENT_ID"],
    authority=f"https://login.microsoftonline.com/{settings['TENANT_ID']}",
    client_credential=settings["CLIENT_SECRET"]
)
token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])["access_token"]
headers = {"Authorization": f"Bearer {token}"}

# 3. サイト内の全ライブラリを走査
site_id = settings['SHAREPOINT_SITE_ID']
drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
drives = requests.get(drives_url, headers=headers).json().get("value", [])

print("--- 検索開始: すべてのライブラリから '_temp' を探します ---")

for drive in drives:
    drive_id = drive['id']
    drive_name = drive['name']

    # 各ライブラリのルート直下を取得
    children_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children"
    items = requests.get(children_url, headers=headers).json().get("value", [])

    for item in items:
        # フォルダかつ名前が _temp か確認
        if item.get('folder') and item['name'] == '_temp':
            print(f"🎉 見つかりました！")
            print(f"   ライブラリ名: {drive_name}")
            print(f"   ドライブID   : {drive_id}")
            print(f"   フォルダID   : {item['id']}")
            print("-" * 30)
            print("【設定用】local.settings.json を以下のように更新してください")
            print(f'"SHAREPOINT_DRIVE_ID": "{drive_id}",')
            print(f'"TEMP_FOLDER_ID": "{item['id']}"')
