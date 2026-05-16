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

@app.route(route="qflowTrigger", auth_level=func.AuthLevel.FUNCTION)
def qflowTrigger(req: func.HttpRequest) -> func.HttpResponse:
    try:
        # 1. リクエスト解析
        req_body = req.get_json()
        pdf_base64 = req_body.get('pdf_content')
        meta = req_body.get('metadata', {})

        # --- フォントの準備（新しいフォルダ構成に対応） ---
        base_path = os.path.dirname(os.path.abspath(__file__))

        # assets_qflow/fonts 内のファイルを参照
        font_path_normal = os.path.join(base_path, "assets_qflow", "fonts", "hiragino_StdNW1.otf")
        font_path_bold   = os.path.join(base_path, "assets_qflow", "fonts", "hiragino_StdNW3.otf")

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
# 3. 全ページへのスタンプ処理
        for page in doc:
            rect = page.rect
            footer_h = 45.36

            # --- A: フォントをこのページに登録する（ここが重要！） ---
            # fontbufferを使って、ページ内で使う「あだ名」を決めます
            page.insert_font(fontname="jp-normal", fontbuffer=font_data_normal)
            page.insert_font(fontname="jp-bold",   fontbuffer=font_data_bold)

            # --- B: 視認性確保（白帯） ---
            footer_rect = fitz.Rect(0, rect.height - footer_h, rect.width, rect.height)
            page.draw_rect(footer_rect, color=(1, 1, 1), fill=(1, 1, 1), fill_opacity=0.9)

            # --- C: スタンプ配置 ---
            # 左：QRコード
            qr_size = 35
            qr_y_pos = rect.height - footer_h + 5
            qr_rect = fitz.Rect(15, qr_y_pos, 15 + qr_size, qr_y_pos + qr_size)
            page.insert_image(qr_rect, stream=qr_bytes)

            # 中央：識別子 (登録した "jp-bold" を使う)
            id_text = f"{meta.get('dept')}-{meta.get('product')}-{meta.get('category')}-{meta.get('year')}-{meta.get('seq')}"
            page.insert_text(
                (65, rect.height - 28),
                id_text,
                fontsize=8,
                fontname="jp-bold",  # ここは fontbuffer ではなく fontname
                color=(0, 0, 0)
            )

            # 中央：属性 (登録した "jp-normal" を使う)
            attr_text = f"改正日: {meta.get('rev_date')} │ 機密レベル: {meta.get('confidential_level')}"
            page.insert_text(
                (65, rect.height - 15),
                attr_text,
                fontsize=7,
                fontname="jp-normal", # ここは fontbuffer ではなく fontname
                color=(0, 0, 0)
            )

            # 右：ページ番号
            page_text = f"Page {page.number + 1} / {doc.page_count}"
            page.insert_text(
                (rect.width - 70, rect.height - 20),
                page_text,
                fontsize=8,
                fontname="jp-normal", # ここは fontbuffer ではなく fontname
                color=(0, 0, 0)
            )

        # 4. セキュリティ設定
        perm = int(fitz.PDF_PERM_PRINT | fitz.PDF_PERM_COPY | fitz.PDF_PERM_ACCESSIBILITY)
        owner_password = os.environ.get('PDF_OWNER_PWD', 'SYSTEM_OWNER_PWD_QFLOW')

        output_buffer = io.BytesIO()
        doc.save(
            output_buffer,
            encryption=fitz.PDF_ENCRYPT_AES_256,
            owner_pw=owner_password,
            permissions=perm,
            garbage=3,
            deflate=True
        )

        return func.HttpResponse(output_buffer.getvalue(), mimetype="application/pdf")

    except Exception as e:
        logging.error(f"システムエラー: {str(e)}")
        return func.HttpResponse(f"Error: {str(e)}", status_code=500)
