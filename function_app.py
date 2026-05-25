import azure.functions as func
import fitz  # PyMuPDF
import segno
import io
import base64
from font_assets import FONT_NORMAL_B64, FONT_BOLD_B64
from logo_assets import LOGO_PNG_B64
import json
import os
import logging
import msal
import requests
import urllib.parse
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

def get_first_string(val) -> str:
    if not val:
        return ""
    if isinstance(val, str):
        return val.strip()
    if isinstance(val, list):
        for item in val:
            if isinstance(item, str) and item.strip():
                return item.strip()
            elif isinstance(item, dict):
                label = item.get('Label') or item.get('Value') or str(item)
                if label and label.strip(): return label.strip()
    if isinstance(val, dict):
        label = val.get('Label') or val.get('Value') or str(val)
        return label.strip() if label else ""
    return str(val).strip()

def get_string_array(val) -> list:
    if not val:
        return []
    if isinstance(val, str):
        return [val.strip()] if val.strip() else []
    if isinstance(val, list):
        extracted = []
        for item in val:
            if isinstance(item, str) and item.strip():
                extracted.append(item.strip())
            elif isinstance(item, dict):
                label = item.get('Label') or item.get('Value') or str(item)
                if label and label.strip(): extracted.append(label.strip())
        return extracted
    if isinstance(val, dict):
        label = val.get('Label') or val.get('Value') or str(val)
        return [label.strip()] if label and label.strip() else []
    return [str(val).strip()] if str(val).strip() else []

# =====================================================================
#  【デザイン拡張版：PDFスタンプ＆暗号化ロジック（確定レイアウト維持）】
# =====================================================================
def process_pdf_stamping(pdf_bytes: bytes, meta: dict, original_filename: str) -> bytes:
    font_data_normal = base64.b64decode(FONT_NORMAL_B64)
    font_data_bold   = base64.b64decode(FONT_BOLD_B64)

    logo_bytes = b""
    if LOGO_PNG_B64 and "YOUR_BASE64" not in LOGO_PNG_B64:
        try:
            logo_bytes = base64.b64decode("".join(LOGO_PNG_B64.split()))
        except Exception as e:
            logging.error(f"ロゴデコード失敗: {e}")

    doc = fitz.open("pdf", io.BytesIO(pdf_bytes))

    qr_url = meta.get('qr_url', 'https://tshldgs.sharepoint.com/sites/QualityAssurance')
    qr_buf = io.BytesIO()
    segno.make_qr(qr_url).save(qr_buf, kind='png', scale=4)
    qr_bytes = qr_buf.getvalue()

    for page in doc:
        rotation = page.rotation
        if rotation != 0: page.set_rotation(0)

        rect = page.rect
        footer_height = 45.0
        y_top_footer = rect.y1 - footer_height

        page.insert_font(fontname="jp-normal", fontbuffer=font_data_normal)
        page.insert_font(fontname="jp-bold",   fontbuffer=font_data_bold)

        page.draw_rect(fitz.Rect(rect.x0, y_top_footer, rect.x1, rect.y1), color=(1, 1, 1), fill=(1, 1, 1), fill_opacity=0.95)

        if logo_bytes:
            logo_rect = fitz.Rect(rect.x0 + 15, y_top_footer + 11.5, rect.x0 + 71, rect.y1 - 11.5)
            page.insert_image(logo_rect, stream=logo_bytes)

        divider_x = rect.x0 + 79
        page.draw_line(fitz.Point(divider_x, y_top_footer + 5), fitz.Point(divider_x, rect.y1 - 5), color=(0.7, 0.7, 0.7), width=0.8)

        y_center = rect.y1 / 2.0
        qr_rect_side = fitz.Rect(rect.x0 + 10, y_center - 17.5, rect.x0 + 45, y_center + 17.5)
        page.insert_image(qr_rect_side, stream=qr_bytes)

        caption_text = "追跡用QR"
        page.insert_text((rect.x0 + 12, y_center + 24), caption_text, fontsize=5.0, fontname="jp-normal", color=(0.5, 0.5, 0.5))

        line1_list = []
        line1_list.append(original_filename)
        dept = get_first_string(meta.get('QS_Department', ''))
        if dept: line1_list.append(dept)
        pg = get_first_string(meta.get('QS_ProductGroup', ''))
        if pg and pg not in ['その他', 'その他 ', '空', '']: line1_list.append(pg)
        line1_text = "  ·  ".join(line1_list)

        control_num = get_first_string(meta.get('QS_ControlNumber', '—'))
        rev_date = meta.get('determined_revision_date', '')
        conf_level = get_first_string(meta.get('QS_ConfLevel', 'Internal'))

        line2_text = f"文書番号: {control_num}  ·  発行日: {rev_date}  ·  開示範囲: {conf_level}  ·  © TOYO SAFETY CO., LTD."
        page_num_text = f"Page {page.number + 1} / {doc.page_count}"

        metadata_start_x = divider_x + 8
        y_line1 = rect.y1 - 26.0
        y_line2 = rect.y1 - 14.0

        page.insert_text((metadata_start_x, y_line1), line1_text, fontsize=7.5, fontname="jp-normal", color=(0, 0, 0))
        page.insert_text((metadata_start_x, y_line2), line2_text, fontsize=6.5, fontname="jp-normal", color=(0.4, 0.4, 0.4))

        page_num_x = rect.x1 - 60
        page.insert_text((page_num_x, y_line2), page_num_text, fontsize=6.5, fontname="jp-normal", color=(0.4, 0.4, 0.4))

        page.draw_line(fitz.Point(rect.x0, y_top_footer), fitz.Point(rect.x1, y_top_footer), color=(0.8, 0.8, 0.8), width=0.5)

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
#  【Azure Functions HTTP Trigger】
# =====================================================================
@app.route(route="StampPDF", auth_level=func.AuthLevel.FUNCTION)
def StampPDF(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("SOW v8.4: StampPDF 文書分類追加・メタデータ完全移植版を開始します。")
    try:
        body = req.get_json()
        distribution_id = body.get("distribution_id")
        temp_file_url = body.get("temp_file_url")

        if not distribution_id or not temp_file_url:
            return func.HttpResponse(json.dumps({"status": "error", "error": "distribution_id or temp_file_url is missing"}), status_code=400)

        token = get_access_token()
        headers = {"Authorization": f"Bearer {token}"}

        site_id = os.environ.get("SHAREPOINT_SITE_ID")
        qlibrary_drive_id = os.environ.get("QLIBRARY_DRIVE_ID")
        qpublished_drive_id = os.environ.get("QPUBLISHED_DRIVE_ID")
        temp_folder_id = os.environ.get("TEMP_FOLDER_ID")
        list_id = os.environ.get("DISTRIBUTION_LOG_LIST_ID")

        if not site_id or not qlibrary_drive_id or not qpublished_drive_id or not temp_folder_id:
            raise Exception("環境変数未設定エラー")

        parsed_url = urllib.parse.urlparse(temp_file_url)
        decoded_path = urllib.parse.unquote(parsed_url.path)
        file_name = decoded_path.split("/")[-1]

        children_url = f"https://graph.microsoft.com/v1.0/drives/{qlibrary_drive_id}/items/{temp_folder_id}/children"
        children_res = requests.get(children_url, headers=headers)
        items = children_res.json().get('value', [])
        item_id = next((i.get('id') for i in items if i.get('name') == file_name), None)
        if not item_id: raise Exception(f"'{file_name}' が見つかりませんでした。")

        convert_url = f"https://graph.microsoft.com/v1.0/drives/{qlibrary_drive_id}/items/{item_id}/content?format=pdf"
        convert_res = requests.get(convert_url, headers=headers, stream=True)
        pdf_bytes = convert_res.content

        meta_url = f"https://graph.microsoft.com/v1.0/drives/{qlibrary_drive_id}/items/{item_id}/listitem?expand=fields"
        meta_res = requests.get(meta_url, headers=headers)
        metadata = meta_res.json().get('fields', {}) if meta_res.status_code == 200 else {}

        determined_date = datetime.now(timezone.utc).strftime("%Y-%m-%d")
        if distribution_id and list_id and site_id:
            log_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items/{distribution_id}?expand=fields"
            log_res = requests.get(log_url, headers=headers)
            if log_res.status_code == 200:
                issue_date = log_res.json().get('fields', {}).get('QS_IssueDate')
                if issue_date: determined_date = issue_date.split("T")[0] if "T" in issue_date else issue_date

        metadata['determined_revision_date'] = determined_date
        base_name, _ = os.path.splitext(file_name)
        output_filename = f"{base_name}.pdf"

        secured_pdf_bytes = process_pdf_stamping(pdf_bytes, metadata, base_name)

        encoded_output_filename = urllib.parse.quote(output_filename)
        upload_url = f"https://graph.microsoft.com/v1.0/drives/{qpublished_drive_id}/root:/{encoded_output_filename}:/content"
        upload_headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/pdf"}
        upload_res = requests.put(upload_url, headers=upload_headers, data=secured_pdf_bytes)

        upload_json = upload_res.json()
        output_file_url = upload_json.get('webUrl')
        published_item_id = upload_json.get('id')

        # 💡 APIのデータ型に合わせて安全な自動移植用ペイロードを構築
        if published_item_id and metadata:
            fields_payload = {}

            # 1. 所管部門 (選択肢)
            dept_val = get_first_string(metadata.get("QS_Department"))
            if dept_val: fields_payload["QS_Department"] = dept_val

            # 2. 商品グループ (複数選択肢) ※型宣言タグをセット
            pg_val = get_string_array(metadata.get("QS_ProductGroup"))
            if pg_val:
                fields_payload["QS_ProductGroup@odata.type"] = "Collection(Edm.String)"
                fields_payload["QS_ProductGroup"] = pg_val

            # 3. 管理番号 (1行テキスト)
            ctrl_val = get_first_string(metadata.get("QS_ControlNumber"))
            if ctrl_val: fields_payload["QS_ControlNumber"] = ctrl_val

            # 4. 機密レベル (選択肢)
            conf_val = get_first_string(metadata.get("QS_ConfLevel"))
            if conf_val: fields_payload["QS_ConfLevel"] = conf_val

            # 5. 改正日 (日付時刻 -> ISO 8601フォーマット)
            rev_date = metadata.get('determined_revision_date')
            if rev_date:
                fields_payload["QS_RevisionDate"] = f"{rev_date}T00:00:00Z" if "T" not in rev_date else rev_date

            # 6. アクセス権限 (選択肢)
            access_val = get_first_string(metadata.get("QS_AccessRights"))
            if access_val: fields_payload["QS_AccessRights"] = access_val

            # 7. QMS要求事項 (選択肢)
            qms_val = get_first_string(metadata.get("QS_QMSRequireNum"))
            if qms_val: fields_payload["QS_QMSRequireNum"] = qms_val

            # 8. 💡【新規追加】文書分類 (選択肢)
            cat_val = get_first_string(metadata.get("QS_DocCategory"))
            if cat_val: fields_payload["QS_DocCategory"] = cat_val

            if fields_payload:
                patch_metadata_url = f"https://graph.microsoft.com/v1.0/drives/{qpublished_drive_id}/items/{published_item_id}/listitem/fields"
                patch_res = requests.patch(patch_metadata_url, headers=headers, json=fields_payload)

                # エラーガード：拒否された場合はログに詳細を出して500エラー停止
                if patch_res.status_code not in [200, 201]:
                    error_msg = f"サイト列の更新に失敗（Graph API拒否）: {patch_res.text}"
                    logging.error(error_msg)
                    raise Exception(error_msg)
                else:
                    logging.info("Q-Publishedへのサイト列メタデータの移植（完全版）が正常に完了しました。")

        delete_url = f"https://graph.microsoft.com/v1.0/drives/{qlibrary_drive_id}/items/{item_id}"
        requests.delete(delete_url, headers=headers)

        result = {
            "status": "success",
            "output_file_url": output_file_url,
            "processed_at": datetime.now(timezone.utc).isoformat()
        }
        return func.HttpResponse(json.dumps(result), mimetype="application/json", status_code=200)

    except Exception as e:
        return func.HttpResponse(json.dumps({"status": "error", "error": str(e)}), status_code=500, mimetype="application/json")
