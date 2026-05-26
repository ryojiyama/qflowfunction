import os
import json
import requests
import msal

# 1. ローカルの local.settings.json から環境変数を強制的に読み込む
def load_local_settings():
    settings_path = os.path.join(os.path.dirname(__file__), 'local.settings.json')
    if os.path.exists(settings_path):
        with open(settings_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
            values = config.get("Values", {})
            for key, val in values.items():
                os.environ[key] = str(val)

# 2. アクセストークン取得 (本番コードと同一ロジック)
def get_access_token():
    tenant_id = os.environ.get("TENANT_ID")
    client_id = os.environ.get("CLIENT_ID")
    client_secret = os.environ.get("CLIENT_SECRET")
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    conf_app = msal.ConfidentialClientApplication(client_id, authority=authority, client_credential=client_secret)
    result = conf_app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" in result:
        return result["access_token"]
    raise Exception(f"トークン取得失敗: {result.get('error_description')}")

def test_inspect_sharepoint_json():
    load_local_settings()
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}

    qlibrary_drive_id = os.environ.get("QLIBRARY_DRIVE_ID")
    temp_folder_id = os.environ.get("TEMP_FOLDER_ID")

    # テストとして、Tempフォルダ内の一番最初のアイテムを取得
    children_url = f"https://graph.microsoft.com/v1.0/drives/{qlibrary_drive_id}/items/{temp_folder_id}/children"
    children_res = requests.get(children_url, headers=headers)
    items = children_res.json().get('value', [])

    if not items:
        print("\n[警告] Tempフォルダにファイルがありません。テスト用にSharePointのTempフォルダに何か1つファイルを配置してください。")
        return

    target_item_id = items[0]['id']
    target_item_name = items[0]['name']
    print(f"\n--- ターゲット検証ファイル: {target_item_name} ---")

    # 💡 パターンA: listitem直下およびfieldsを展開して取得するURL
    meta_url = f"https://graph.microsoft.com/v1.0/drives/{qlibrary_drive_id}/items/{target_item_id}/listitem?expand=fields"
    meta_res = requests.get(meta_url, headers=headers)
    res_json = meta_res.json()

    print("\n==================【パターンA: listitem 全体構造】==================")
    # バージョンに関するキーワードが直下に含まれているか検索
    version_keys_in_root = {k: v for k, v in res_json.items() if "version" in k.lower()}
    print(f"直下のプロパティ内での検索結果: {version_keys_in_root}")

    print("\n==================【パターンA: fields 内部の構造】==================")
    fields = res_json.get("fields", {})
    # fieldsのキーの中に "version" や "UIVersion" が含まれているか検索
    version_keys_in_fields = {k: v for k, v in fields.items() if "version" in k.lower() or "uiv" in k.lower()}
    print(f"fields内での検索結果: {version_keys_in_fields}")

    # 💡 パターンB: ドライブアイテム（driveItem）自体のプロパティも確認
    item_url = f"https://graph.microsoft.com/v1.0/drives/{qlibrary_drive_id}/items/{target_item_id}"
    item_res = requests.get(item_url, headers=headers)
    item_json = item_res.json()

    print("\n==================【パターンB: driveItem の構造】==================")
    item_version = item_json.get("version") # driveItem標準のversionプロパティ
    print(f"driveItem.version の値: {item_version}")

if __name__ == "__main__":
    test_inspect_sharepoint_json()
