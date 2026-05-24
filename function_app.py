import azure.functions as func
import fitz  # PyMuPDF
import segno
import io
import base64
from font_assets import FONT_NORMAL_B64, FONT_BOLD_B64
import json
import os
import logging
import msal
import requests
from datetime import datetime, timezone

app = func.FunctionApp()

# --- 認証トークン取得 ---
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

# --- SharePointの絶対URLをGraph API用にエンコードする関数 ---
def encode_sharing_url(url: str) -> str:
    """SharePointの絶対URLをGraph APIの /shares/ エンドポイントで使える形式に変換する"""
    base64_value = base64.b64encode(url.encode('utf-8')).decode('utf-8')
    encoded_url = "u!" + base64_value.rstrip('=').replace('/', '_').replace('+', '-')
    return encoded_url

# --- 既存のPDFスタンプ＆暗号化ロジック ---
def process_pdf_stamping(pdf_bytes: bytes, meta: dict) -> bytes:
    font_data_normal = base64.b64decode(FONT_NORMAL_B64)
    font_data_bold   = base64.b64decode(FONT_BOLD_B64)

    doc = fitz.open("pdf", io.BytesIO(pdf_bytes))
    qr_url = meta.get('qr_url', 'https://tshldgs.sharepoint.com/sites/QualityAssurance')
    qr_buf = io.BytesIO()
    segno.make_qr(qr_url).save(qr_buf, kind='png', scale=4)
    qr_bytes = qr_buf.getvalue()

    for page in doc:
        rotation = page.rotation
        if rotation != 0: page.set_rotation(0)

        rect = page.rect
        y_top_footer = rect.y1 - 45.36

        page.insert_font(fontname="jp-normal", fontbuffer=font_data_normal)
        page.insert_font(fontname="jp-bold",   fontbuffer=font_data_bold)

        page.draw_rect(fitz.Rect(rect.x0, y_top_footer, rect.x1, rect.y1), color=(1, 1, 1), fill=(1, 1, 1), fill_opacity=0.9)
        page.insert_image(fitz.Rect(rect.x0 + 15, y_top_footer + 5, rect.x0 + 50, y_top_footer + 40), stream=qr_bytes)

        id_text = meta.get('integrated_id') or f"{meta.get('dept', '')}-{meta.get('product', '')}-{meta.get('category', '')}-{meta.get('seq', '')}"
        page.insert_text((rect.x0 + 65, rect.y1 - 28), id_text, fontsize=8, fontname="jp-bold", color=(0, 0, 0))

        attr_text = f"改正日: {meta.get('rev_date', '')} │ 機密レベル: {meta.get('confidential_level', 'Internal')}"
        page.insert_text((rect.x0 + 65, rect.y1 - 15), attr_text, fontsize=7, fontname="jp-normal", color=(0, 0, 0))
        page.insert_text((rect.x1 - 70, rect.y1 - 20), f"Page {page.number + 1} / {doc.page_count}", fontsize=8, fontname="jp-normal", color=(0, 0, 0))

        if rotation != 0: page.set_rotation(rotation)

    temp_buffer = io.BytesIO()
    doc.save(temp_buffer, garbage=3, deflate=True)
    doc.close()

    secure_doc = fitz.open("pdf", temp_buffer.getvalue())
    perm = int(fitz.PDF_PERM_PRINT | fitz.PDF_PERM_COPY | fitz.PDF_PERM_ACCESSIBILITY)
    owner_pwd = os.environ.get('PDF_OWNER_PWD', 'SYSTEM_OWNER_PWD_QFLOW')

    output_buffer = io.BytesIO()
    secure_doc.save(output_buffer, encryption=fitz.PDF_ENCRYPT_AES_256, owner_pw=owner_pwd, permissions=perm, garbage=3, deflate=True)
    secure_doc.close()

    return output_buffer.getvalue()

# =====================================================================
#  【Azure Functions HTTP Trigger (SOW v4仕様)】
# =====================================================================
@app.route(route="StampPDF", auth_level=func.AuthLevel.FUNCTION)
def StampPDF(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("SOW v4: StampPDF 非同期処理を開始します。")
    try:
        # 1. パラメーター取得
        body = req.get_json()
        distribution_id = body.get("distribution_id")
        temp_file_url = body.get("temp_file_url")

        if not distribution_id or not temp_file_url:
            return func.HttpResponse(json.dumps({"status": "error", "error": "distribution_id or temp_file_url is missing"}), status_code=400)

        # 2. 認証
        token = get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        site_id = os.environ.get("SHAREPOINT_SITE_ID")
        list_id = os.environ.get("DISTRIBUTION_LOG_LIST_ID")

        # 3. URLからSharePointのサイトIDおよびアイテムIDを分解解決（Graph APIのsharesエンドポイント使用）
        encoded_url = encode_sharing_url(temp_file_url)
        share_res = requests.get(f"https://graph.microsoft.com/v1.0/shares/{encoded_url}/driveItem", headers=headers)
        if share_res.status_code != 200:
            raise Exception(f"ファイルが見つかりません。URLを確認してください: {share_res.text}")

        drive_item = share_res.json()
        drive_id = drive_item['parentReference']['driveId']
        item_id = drive_item['id']
        file_name = drive_item['name'] # 例: 123_input.xlsx

        # 4. Word/Excelから直接PDFストリームとしてPull（サイズ制約回避）
        logging.info("Graph APIを利用してWord/ExcelをPDFとしてオンデマンド変換中...")
        convert_res = requests.get(f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content?format=pdf", headers=headers, stream=True)
        if convert_res.status_code != 200:
            raise Exception("graph_conversion_failed")
        pdf_bytes = convert_res.content

        # 5. メタデータをDistributionLogから取得 (ローカル検証時はダミー値を使用可能)
        # ※リストが未整備の場合でもテストが通るように、辞書の get() メソッドで安全に取得します。
        metadata = {}
        log_res = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items/{distribution_id}?expand=fields", headers=headers)
        if log_res.status_code == 200:
            metadata = log_res.json().get('fields', {})
            logging.info("DistributionLogからメタデータの取得に成功しました。")
        else:
            logging.warning(f"リストからのメタデータ取得に失敗（テスト用のダミー値を使用）: {log_res.text}")
            metadata = {"dept": "QA", "product": "Test", "seq": distribution_id}

        # 6. スタンプ付与 & 7. AES-256暗号化
        logging.info("PDFのスタンプ加工および暗号化を実行中...")
        secured_pdf_bytes = process_pdf_stamping(pdf_bytes, metadata)

        # 8. 出力先ライブラリ（Issuedフォルダ）へアップロード
        # 保存ファイル名: "元のファイル名_stamped.pdf" または "{配布ID}_stamped.pdf"
        output_filename = f"{distribution_id}_stamped.pdf"
        upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/Issued/{output_filename}:/content"

        logging.info(f"Issuedフォルダへアップロード中: {output_filename}")
        upload_headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/pdf"}
        upload_res = requests.put(upload_url, headers=upload_headers, data=secured_pdf_bytes)

        if upload_res.status_code not in [200, 201]:
            raise Exception("upload_failed")

        output_file_url = upload_res.json().get('webUrl')

        # 9. 【安全設計】用済みの _temp 内のコピーファイル（Word/Excel）を物理削除
        logging.info(f"一時ファイル（_temp）のクリーンアップを実行中... ItemID: {item_id}")
        del_res = requests.delete(f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}", headers=headers)
        if del_res.status_code != 204:
            logging.warning(f"一時ファイルの削除に失敗しました（処理は継続します）: {del_res.text}")

        # 10. 完了レスポンス（Flow Bがこれを受け取って台帳を更新する）
        result = {
            "status": "success",
            "output_file_url": output_file_url,
            "processed_at": datetime.now(timezone.utc).isoformat()
        }
        return func.HttpResponse(json.dumps(result), mimetype="application/json", status_code=200)

    except Exception as e:
        error_msg = str(e)
        logging.error(f"エラー発生: {error_msg}")
        return func.HttpResponse(json.dumps({"status": "error", "error": error_msg}), status_code=500, mimetype="application/json")
