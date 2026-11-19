import streamlit as st
import streamlit.components.v1 as components
import base64
import os

# ページの基本設定
st.set_page_config(page_title="エクセルコピーツール", layout="centered")

st.title("📄 エクセル完全コピーツール")
st.write("元のエクセルをドラッグ＆ドロップすると、書式・数式・複数タブをそのまま維持してコピーを作成します。")

# 1. ファイルのアップロード
uploaded_file = st.file_uploader(
    "コピー元のエクセルファイルを選択（またはここにドロップ）", 
    type=["xlsx", "xlsm", "xls"]
)

if uploaded_file is not None:
    original_name, ext = os.path.splitext(uploaded_file.name)
    
    st.divider()
    st.subheader("⚙️ コピー設定")
    
    # 2. 新しいファイル名の指定
    new_filename = st.text_input(
        "新しいファイル名（拡張子は自動で付きます）", 
        value=f"{original_name}_コピー"
    )
    
    target_filename = f"{new_filename}{ext}"
    
    # 3. データをBase64文字列に変換（ブラウザ内で直接ファイル化するため）
    file_bytes = uploaded_file.getvalue()
    b64_data = base64.b64encode(file_bytes).decode()
    
    # 4. ブラウザ内で直接ダウンロードを実行するボタン（セキュリティ制限回避）
    download_html = f"""
    <html>
    <head>
    <style>
        body {{ margin: 0; padding: 0; font-family: sans-serif; }}
        .btn {{
            background-color: #FF4B4B;
            color: white;
            border: none;
            padding: 10px 20px;
            font-size: 16px;
            font-weight: bold;
            border-radius: 8px;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 8px;
        }}
        .btn:hover {{ background-color: #FF3333; }}
    </style>
    </head>
    <body>
        <button class="btn" onclick="downloadExcel()">📥 コピーを作成してダウンロード</button>
        <script>
        function downloadExcel() {{
            const b64 = "{b64_data}";
            const byteCharacters = atob(b64);
            const byteNumbers = new Array(byteCharacters.length);
            for (let i = 0; i < byteCharacters.length; i++) {{
                byteNumbers[i] = byteCharacters.charCodeAt(i);
            }}
            const byteArray = new Uint8Array(byteNumbers);
            const blob = new Blob([byteArray], {{type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}});
            
            const link = document.createElement("a");
            link.href = URL.createObjectURL(blob);
            link.download = "{target_filename}";
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }}
        </script>
    </body>
    </html>
    """
    components.html(download_html, height=60)
