import msal, requests, json

with open("local.settings.json", "r") as f:
    settings = json.load(f)["Values"]

app = msal.ConfidentialClientApplication(
    settings["CLIENT_ID"],
    authority=f"https://login.microsoftonline.com/{settings['TENANT_ID']}",
    client_credential=settings["CLIENT_SECRET"]
)
token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])["access_token"]
headers = {"Authorization": f"Bearer {token}"}

site_id = settings['SHAREPOINT_SITE_ID']
url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
drives = requests.get(url, headers=headers).json().get("value", [])

print("--- サイト内のドキュメントライブラリ一覧 ---")
for drive in drives:
    # driveTypeがdocumentLibraryのものだけを表示
    if drive.get('driveType') == 'documentLibrary':
        print(f"名前: {drive['name']} | ID: {drive['id']}")
