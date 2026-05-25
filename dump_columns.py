import os
import json
import requests
import msal

# --- 1. local.settings.json から環境変数を強制ロード ---
def load_local_settings():
    settings_path = "local.settings.json"
    if os.path.exists(settings_path):
        with open(settings_path, "r", encoding="utf-8") as f:
            data = json.load(f)
            values = data.get("Values", {})
            for k, v in values.items():
                os.environ[k] = str(v)
        print("💡 local.settings.json から環境変数を読み込みました。")
    else:
        print("⚠️ local.settings.json が見つかりません。現在の環境変数を使用します。")

# --- 2. 認証トークン取得 ---
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

# --- 3. 特定のエンドポイントから列一覧を取得して表示する関数 ---
def dump_columns(url, headers, title_name):
    print(f"\n==================================================")
    print(f" 🔍 {title_name} の列定義一覧")
    print(f"==================================================")

    res = requests.get(url, headers=headers)
    if res.status_code != 200:
        print(f"❌ 取得失敗: {res.status_code} - {res.text}")
        return

    columns = res.json().get("value", [])
    # 読みやすさのために表示名でソート
    columns.sort(key=lambda x: x.get("displayName", ""))

    print(f"{'【表示名 (画面上の名前)】':<30} ➔ {'【内部名 (コードに書くべき正解のキー)】'}")
    print("-" * 80)

    for col in columns:
        display_name = col.get("displayName", "N/A")
        internal_name = col.get("name", "N/A")
        # システム固有の標準列（idやContentTypeなど）は除外せず全て出力
        print(f"{display_name:<30} ➔ {internal_name}")

# --- メイン処理 ---
def main():
    load_local_settings()

    try:
        token = get_access_token()
        headers = {"Authorization": f"Bearer {token}"}

        site_id = os.environ.get("SHAREPOINT_SITE_ID")
        list_id = os.environ.get("DISTRIBUTION_LOG_LIST_ID")
        qlibrary_id = os.environ.get("QLIBRARY_DRIVE_ID")
        qpublished_id = os.environ.get("QPUBLISHED_DRIVE_ID")

        # ① DistributionLog リストの列を取得
        if site_id and list_id:
            list_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/columns"
            dump_columns(list_url, headers, "1. DistributionLog (管理リスト)")

        # ② Q-Library ドキュメントライブラリの列を取得
        if qlibrary_id:
            qlib_url = f"https://graph.microsoft.com/v1.0/drives/{qlibrary_id}/list/columns"
            dump_columns(qlib_url, headers, "2. Q-Library (_tempの元ライブラリ)")

        # ③ Q-Published ドキュメントライブラリの列を取得
        if qpublished_id:
            qpub_url = f"https://graph.microsoft.com/v1.0/drives/{qpublished_id}/list/columns"
            dump_columns(qpub_url, headers, "3. Q-Published (配布先ライブラリ)")

    except Exception as e:
        print(f"🚨 致命的エラー: {e}")

if __name__ == "__main__":
    main()
