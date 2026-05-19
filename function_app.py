import azure.functions as func
import fitz  # PyMuPDF
import segno
import io
import base64
from font_assets import FONT_NORMAL_B64, FONT_BOLD_B64
import json
import os
import logging

app = func.FunctionApp()

@app.route(route="StampPDF", auth_level=func.AuthLevel.FUNCTION)
def StampPDF(req: func.HttpRequest) -> func.HttpResponse:
    try:
        # 1. リクエスト解析
        req_body = req.get_json()
        pdf_base64 = req_body.get('pdf_content')
        meta = req_body.get('metadata', {})

        # --- フォントの準備（ファイルシステムを一切使わない方式） ---
        try:
            # 埋め込まれた文字列をメモリ上でデコード
            font_data_normal = base64.b64decode(FONT_NORMAL_B64)
            font_data_bold   = base64.b64decode(FONT_BOLD_B64)
            logging.info("埋め込みフォントの展開に成功しました。")
        except Exception as e:
            logging.error(f"フォント展開失敗: {e}")
            return func.HttpResponse("Font Data Error", status_code=500)

        # 2. PDFとQRコードの準備
        doc = fitz.open("pdf", io.BytesIO(base64.b64decode(pdf_base64)))

        # QRコード生成
        qr_url = meta.get('qr_url', 'https://tshldgs.sharepoint.com/sites/QualityAssurance')
        qr_buf = io.BytesIO()
        segno.make_qr(qr_url).save(qr_buf, kind='png', scale=4)
        qr_bytes = qr_buf.getvalue()

        # 3. 全ページへのスタンプ処理
        for page in doc:
            # ページが持つ固有の回転角度（90, 180, 270など）を取得
            rotation = page.rotation
            if rotation != 0:
                page.set_rotation(0)  # 描画位置の破綻を防ぐため、一時的に0度にリセット

            rect = page.rect
            footer_h = 45.36

            # 絶対座標の定義（原点ズレに対応）
            x_left = rect.x0
            x_right = rect.x1
            y_bottom = rect.y1
            y_top_footer = y_bottom - footer_h

            # --- A: フォントをこのページに登録する ---
            page.insert_font(fontname="jp-normal", fontbuffer=font_data_normal)
            page.insert_font(fontname="jp-bold",   fontbuffer=font_data_bold)

            # --- B: 視認性確保（白帯）を絶対座標で描画 ---
            footer_rect = fitz.Rect(x_left, y_top_footer, x_right, y_bottom)
            page.draw_rect(footer_rect, color=(1, 1, 1), fill=(1, 1, 1), fill_opacity=0.9)

            # --- C: スタンプ配置 ---
            # 左：QRコード
            qr_size = 35
            qr_y_pos = y_top_footer + 5
            qr_rect = fitz.Rect(x_left + 15, qr_y_pos, x_left + 15 + qr_size, qr_y_pos + qr_size)
            page.insert_image(qr_rect, stream=qr_bytes)

            # 中央：識別子
            if meta.get('integrated_id'):
                id_text = meta.get('integrated_id')
            else:
                id_text = f"{meta.get('dept')}-{meta.get('product')}-{meta.get('category')}-{meta.get('year')}-{meta.get('seq')}"
            page.insert_text(
                (x_left + 65, y_bottom - 28),
                id_text,
                fontsize=8,
                fontname="jp-bold",
                color=(0, 0, 0)
            )

            # 中央：属性
            attr_text = f"改正日: {meta.get('rev_date')} │ 機密レベル: {meta.get('confidential_level')}"
            page.insert_text(
                (x_left + 65, y_bottom - 15),
                attr_text,
                fontsize=7,
                fontname="jp-normal",
                color=(0, 0, 0)
            )

            # 右：ページ番号
            page_text = f"Page {page.number + 1} / {doc.page_count}"
            page.insert_text(
                (x_right - 70, y_bottom - 20),
                page_text,
                fontsize=8,
                fontname="jp-normal",
                color=(0, 0, 0)
            )

            # すべての描画完了後、元の回転角度に戻す
            if rotation != 0:
                page.set_rotation(rotation)
        # 4. 描画内容を確定させるため、一度暗号化なしでメモリ上に仮保存する
        temp_buffer = io.BytesIO()
        doc.save(temp_buffer, garbage=3, deflate=True)
        doc.close()

        # 5. 描画が確定したPDFデータを再度開き、セキュリティ設定のみを適用する
        secure_doc = fitz.open("pdf", temp_buffer.getvalue())

        perm = int(fitz.PDF_PERM_PRINT | fitz.PDF_PERM_COPY | fitz.PDF_PERM_ACCESSIBILITY)
        owner_password = os.environ.get('PDF_OWNER_PWD', 'SYSTEM_OWNER_PWD_QFLOW')

        output_buffer = io.BytesIO()
        secure_doc.save(
            output_buffer,
            encryption=fitz.PDF_ENCRYPT_AES_256,
            owner_pw=owner_password,
            permissions=perm,
            garbage=3,
            deflate=True
        )
        secure_doc.close()

        return func.HttpResponse(output_buffer.getvalue(), mimetype="application/pdf")

    except Exception as e:
        logging.error(f"システムエラー: {str(e)}")
        return func.HttpResponse(f"Error: {str(e)}", status_code=500)
