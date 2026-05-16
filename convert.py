import base64

def create_assets():
    # 既存のフォントファイルを読み込んでBase64化
    with open("assets_qflow/fonts/hiragino_StdNW1.otf", "rb") as f:
        normal = base64.b64encode(f.read()).decode('utf-8')
    with open("assets_qflow/fonts/hiragino_StdNW3.otf", "rb") as f:
        bold = base64.b64encode(f.read()).decode('utf-8')

    # font_assets.py を自動作成
    with open("font_assets.py", "w") as f:
        f.write(f'FONT_NORMAL_B64 = "{normal}"\n')
        f.write(f'FONT_BOLD_B64 = "{bold}"\n')
    print("font_assets.py の作成が完了しました！")

if __name__ == "__main__":
    create_assets()
