import streamlit as st
import streamlit.components.v1 as components
import openpyxl
import io
import base64
import os
import re
import datetime
import calendar

# ページの基本設定
st.set_page_config(page_title="エクセルコピーツール（全パターン対応版）", layout="centered")

st.title("📄 エクセル白紙化コピーツール")
st.write("レイアウト・色・計算式・見出しはそのまま残し、**手入力された「数値」を自動でクリア**します。")

# 1. ファイルのアップロード
uploaded_file = st.file_uploader(
    "コピー元のエクセルファイルを選択（またはここにドロップ）", 
    type=["xlsx", "xlsm"]
)

if uploaded_file is not None:
    original_name, ext = os.path.splitext(uploaded_file.name)
    
    st.divider()
    st.subheader("⚙️ オプション設定")
    
    # 年更新の切り替えスイッチ
    update_year = st.checkbox("📅 カレンダー・日付・曜日を翌年（2026 → 2027）に更新する", value=True)
    
    # ファイル名の自動提案（年更新ONなら2026→2027も反映、OFFなら月だけ翌月提案）
    def suggest_filename(name, is_next_year):
        new_name = name
        if is_next_year:
            new_name = new_name.replace("2026", "2027")
        else:
            match = re.search(r'(\d{1,2})月', new_name)
            if match:
                month = int(match.group(1))
                next_month = 1 if month == 12 else month + 1
                new_name = re.sub(r'(\d{1,2})月', f'{next_month}月', new_name, count=1)
        if new_name == name:
            return f"{name}_コピー"
        return new_name

    default_new_name = suggest_filename(original_name, update_year)
    
    # 2. openpyxlでエクセルを読み込み
    keep_vba = (ext.lower() == ".xlsm")
    wb = openpyxl.load_workbook(uploaded_file, keep_vba=keep_vba)
    
    cleared_count = 0
    updated_date_count = 0
    
    # 曜日変換用の辞書
    WEEKDAYS_JP = ["月", "火", "水", "木", "金", "土", "日"]
    
    def format_weekday(old_str, new_wd_index):
        """元の曜日表記スタイル（月 / (月) / 月曜日 など）を保ったまま新しい曜日に書き換える"""
        new_wd = WEEKDAYS_JP[new_wd_index]
        # 括弧つき (月) や （月）
        if re.fullmatch(r'[\(（][月火水木金土日][\)）]', old_str.strip()):
            return re.sub(r'[月火水木金土日]', new_wd, old_str)
        # 〇〇曜日
        if re.fullmatch(r'[月火水木金土日]曜日', old_str.strip()):
            return f"{new_wd}曜日"
        # 1文字だけ 月, 火...
        if old_str.strip() in WEEKDAYS_JP:
            return new_wd
        return None

    # 全シートを順番に処理
    for sheet in wb.worksheets:
        # シート名に「2026」があれば「2027」に変更
        if update_year and "2026" in sheet.title:
            sheet.title = sheet.title.replace("2026", "2027")
            
        # シート全体から「何月」のシートかを推測（1〜31日カレンダーの曜日計算用）
        sheet_month = None
        m_title = re.search(r'(\d{1,2})月', sheet.title)
        if m_title:
            sheet_month = int(m_title.group(1))
            
        # まずシート内の年・月情報をスキャンしつつ、日付と文字列を更新
        for row in sheet.iter_rows():
            for cell in row:
                val = cell.value
                if val is None:
                    continue
                
                # シート内の見出しから「〇月」を探す（シート名に月がない場合用）
                if sheet_month is None and isinstance(val, str):
                    m_cell = re.search(r'202[67]年\s*(\d{1,2})月', val)
                    if m_cell:
                        sheet_month = int(m_cell.group(1))

                # 計算式はそのまま残す
                if isinstance(val, str) and val.startswith("="):
                    continue
                
                # 【パターン1】日付・時刻データ（シリアル値）の年を+1する
                if isinstance(val, (datetime.datetime, datetime.date)):
                    if update_year and val.year == 2026:
                        try:
                            cell.value = val.replace(year=2027)
                        except ValueError:
                            # 2月29日対策
                            cell.value = val.replace(year=2027, day=28)
                        updated_date_count += 1
                    continue
                
                if isinstance(val, datetime.time) or isinstance(val, bool):
                    continue
                
                # 【パターン2】文字列内の「2026」を「2027」に置換
                if isinstance(val, str):
                    if update_year and "2026" in val:
                        cell.value = val.replace("2026", "2027")
                        updated_date_count += 1
                    continue
                
                # 【数値の処理】
                if isinstance(val, (int, float)):
                    # 年号として「2026」が単独で入っているセルは「2027」に更新して残す
                    if update_year and val == 2026:
                        cell.value = 2027
                        updated_date_count += 1
                        continue
                    
                    # 【パターン3】「1〜31」の数字の隣（右または下）に曜日文字がある場合は「日付番号」とみなして残し、曜日を2027年に更新
                    is_calendar_day = False
                    if isinstance(val, int) and 1 <= val <= 31:
                        # 右隣と下のセルをチェック
                        neighbors = [
                            sheet.cell(row=cell.row, column=cell.column + 1),
                            sheet.cell(row=cell.row + 1, column=cell.column)
                        ]
                        for nb in neighbors:
                            if isinstance(nb.value, str) and re.search(r'[月火水木金土日]', nb.value.strip()) and len(nb.value.strip()) <= 4:
                                is_calendar_day = True
                                if update_year and sheet_month and 1 <= sheet_month <= 12:
                                    max_day = calendar.monthrange(2027, sheet_month)[1]
                                    if val <= max_day:
                                        new_wd_idx = datetime.date(2027, sheet_month, val).weekday()
                                        formatted = format_weekday(nb.value, new_wd_idx)
                                        if formatted:
                                            nb.value = formatted
                                            updated_date_count += 1
                    
                    if is_calendar_day:
                        continue
                    
                    # それ以外の純粋な手入力数値（実績・金額など）はクリア
                    cell.value = None
                    cleared_count += 1

    # 処理後のエクセルをメモリ上に保存
    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    processed_bytes = output_buffer.getvalue()
    b64_data = base64.b64encode(processed_bytes).decode()
    
    # 結果レポート表示
    if update_year:
        st.success(f"✅ 処理完了：手入力数値 **{cleared_count} 箇所** をクリアし、日付・年号・曜日を **{updated_date_count} 箇所** 更新（2027年版）しました！")
    else:
        st.success(f"✅ 処理完了：手入力数値 **{cleared_count} 箇所** をクリアしました！")
    
    st.divider()
    
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
        
        <button class="btn" onclick="downloadExcel()">📥 この名前でコピーを作成してダウンロード</button>

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
