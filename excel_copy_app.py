import streamlit as st
import streamlit.components.v1 as components
import base64
import os
import re

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
    
    # ファイル名に「〇月」が含まれていたら、自動で翌月を提案する
    def suggest_next_month(name):
        match = re.search(r'(\d{1,2})月', name)
        if match:
            month = int(match.group(1))
            next_month = 1 if month == 12 else month + 1
            return re.sub(r'(\d{1,2})月', f'{next_month}月', name, count=1)
        return f"{name}_コピー"

    default_new_name = suggest_next_month(original_name)
    
    st.divider()
    st.subheader("⚙️ コピー設定")
    
    # 2. データをBase64文字列に変換
    file_bytes = uploaded_file.getvalue()
    b64_data = base64.b64encode(file_bytes).decode()
    
    # 3. 入力欄とダウンロードボタンを一体化（Enter不要・1クリックで即リネーム＆保存）
    integrated_html = f"""
    <html>
    <head>
    <style>
        body {{
            margin: 0;
            padding: 4px 2px;
            font-family: "Source Sans Pro", sans-serif;
            color: #31333F;
        }}
        .label {{
            font-size: 14px;
            margin-bottom: 8px;
            display: block;
        }}
        .input-group {{
            display: flex;
            align-items: center;
            margin-bottom: 20px;
        }}
        .text-input {{
            flex: 1;
            padding: 10px 12px;
            font-size: 16px;
            border: 1px solid #D3D3D3;
            border-radius: 8px;
            background-color: #F0F2F6;
            color: #31333F;
            outline: none;
        }}
        .text-input:focus {{
            border-color: #FF4B4B;
            background-color: #FFFFFF;
        }}
        .ext-badge {{
            margin-left: 8px;
            font-size: 16px;
            font-weight: bold;
            color: #555;
        }}
        .btn {{
            background-color: #FF4B4B;
            color: white;
            border: none;
            padding: 12px 24px;
            font-size: 16px;
            font-weight: bold;
            border-radius: 8px;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 8px;
            transition: background-color 0.2s;
        }}
        .btn:hover {{
            background-color: #FF3333;
        }}
    </style>
    </head>
    <body>
        <label class="label">新しいファイル名（そのまま書き換えて下のボタンを押してください）</label>
        <div class="input-group">
            <input type="text" id="filenameInput" class="text-input" value="{default_new_name}">
            <span class="ext-badge">{ext}</span>
        </div>
        
        <button class="btn" onclick="downloadExcel()">📥 この名前でコピーを作成してダウンロード</button>

        <script>
        function downloadExcel() {{
            // ボタンを押した瞬間の入力欄の文字を直接取得
            let name = document.getElementById("filenameInput").value.trim();
            if (!name) {{
                name = "{default_new_name}";
            }}
            const targetFilename = name + "{ext}";

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
            link.download = targetFilename;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }}
        </script>
    </body>
    </html>
    """
    components.html(integrated_html, height=150)
