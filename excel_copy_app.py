import streamlit as st
import streamlit.components.v1 as components
import openpyxl
import io
import base64
import os
import re
import datetime

# ページの基本設定
st.set_page_config(page_title="エクセルコピーツール（パターンA実験）", layout="centered")

st.title("📄 エクセル白紙化コピーツール（実験版）")
st.write("レイアウト・色・計算式・見出し文字はそのまま残し、**手入力された「数値」だけを自動でクリア**します。")

# 1. ファイルのアップロード
uploaded_file = st.file_uploader(
    "コピー元のエクセルファイルを選択（またはここにドロップ）", 
    type=["xlsx", "xlsm"]
)

if uploaded_file is not None:
    original_name, ext = os.path.splitext(uploaded_file.name)
    
    # 翌月自動提案
    def suggest_next_month(name):
        match = re.search(r'(\d{1,2})月', name)
        if match:
            month = int(match.group(1))
            next_month = 1 if month == 12 else month + 1
            return re.sub(r'(\d{1,2})月', f'{next_month}月', name, count=1)
        return f"{name}_白紙コピー"

    default_new_name = suggest_next_month(original_name)
    
    # 2. openpyxlでエクセルを読み込み、パターンAの処理を実行
    # マクロ付き(.xlsm)にも対応できるようkeep_vbaを設定
    keep_vba = (ext.lower() == ".xlsm")
    wb = openpyxl.load_workbook(uploaded_file, keep_vba=keep_vba)
    
    cleared_count = 0
    
    # 全シートを順番にチェック
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                val = cell.value
                # 空欄はスキップ
                if val is None:
                    continue
                # 計算式（'='で始まる文字列）は残す
                if isinstance(val, str) and val.startswith("="):
                    continue
                # 日付・時刻データは残す
                if isinstance(val, (datetime.datetime, datetime.date, datetime.time)):
                    continue
                # ブール値（True/False）は残す
                if isinstance(val, bool):
                    continue
                # 手入力された数値（整数・小数）だけをクリアする
                if isinstance(val, (int, float)):
                    cell.value = None
                    cleared_count += 1
                    
    # 処理後のエクセルをメモリ上に保存
    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    processed_bytes = output_buffer.getvalue()
    b64_data = base64.b64encode(processed_bytes).decode()
    
    st.success(f"✅ 処理完了：全 {len(wb.worksheets)} シートから、手入力された数値 **{cleared_count} 箇所** をクリアしました！（計算式と文字は維持されています）")
    
    st.divider()
    st.subheader("⚙️ コピー設定")
    
    # 3. 入力欄とダウンロードボタンの一体化UI
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
        
        <button class="btn" onclick="downloadExcel()">📥 数値をクリアしたコピーをダウンロード</button>

        <script>
        function downloadExcel() {{
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
