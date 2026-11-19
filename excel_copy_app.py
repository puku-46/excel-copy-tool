import streamlit as st
import streamlit.components.v1 as components
import openpyxl
import io
import base64
import os
import re
import datetime
import calendar

st.set_page_config(page_title="エクセル白紙化コピーツール", layout="centered")

st.title("📄 エクセル白紙化コピーツール")
st.write("アップロードしたエクセルを解析し、変更対象を確認してからコピーを作成します。")

uploaded_file = st.file_uploader("コピー元のエクセルファイルを選択（またはここにドロップ）", type=["xlsx", "xlsm"])

if uploaded_file is not None:
    original_name, ext = os.path.splitext(uploaded_file.name)
    keep_vba = (ext.lower() == ".xlsm")
    wb_scan = openpyxl.load_workbook(uploaded_file, keep_vba=keep_vba)

    found_years = set()
    found_samples = []

    for sheet in wb_scan.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                val = cell.value
                if isinstance(val, str) and not val.startswith("="):
                    matches = re.findall(r'(20\d{2})(年度|年)?', val)
                    for y, suffix in matches:
                        found_years.add(y)
                        label = f"{y}{suffix}"
                        if label not in found_samples and len(found_samples) < 6:
                            found_samples.append(label)
                elif isinstance(val, (datetime.datetime, datetime.date)):
                    found_years.add(str(val.year))
                    label = val.strftime("%Y/%m/%d")
                    if label not in found_samples and len(found_samples) < 6:
                        found_samples.append(label)
                elif isinstance(val, int) and 2020 <= val <= 2035:
                    found_years.add(str(val))

    st.divider()
    st.subheader("🔍 1. 事前チェック：検出された年号・日付")

    if found_years:
        sorted_years = sorted(list(found_years))
        st.info(f"シート内から年号 **{', '.join(sorted_years)}** が見つかりました。（例: {', '.join(found_samples)}）")
    else:
        st.info("シート内から特定の年号（202X）は検出されませんでした。")

    year_mode = st.radio(
        "📅 年号・日付の変更方式を選んでください：",
        options=[
            "すべての年号・日付を ＋1年 スライドする（例: 2025→2026、2026→2027、2027→2028）",
            "特定の年だけを指定して書き換える",
            "年号・日付は変更しない（そのまま残す）"
        ],
        index=0
    )

    target_old_year = ""
    target_new_year = ""
    if year_mode == "特定の年だけを指定して書き換える":
        col1, col2 = st.columns(2)
        with col1:
            target_old_year = st.text_input("変更前の年（半角数字）", value="2026")
        with col2:
            target_new_year = st.text_input("変更後の年（半角数字）", value="2027")

    st.divider()
    st.subheader("🧹 2. 数値クリアのルール設定")

    col_opt1, col_opt2 = st.columns(2)
    with col_opt1:
        keep_decimals = st.checkbox("小数点を含む数値（単価など）は残す", value=True)
    with col_opt2:
        protect_keywords = st.text_input("残したい項目のキーワード（カンマ区切り）", value="基本料金,単価,単年割,割合")

    keyword_list = [k.strip() for k in protect_keywords.split(",") if k.strip()]

    uploaded_file.seek(0)
    wb = openpyxl.load_workbook(uploaded_file, keep_vba=keep_vba)

    cleared_count = 0
    updated_date_count = 0
    WEEKDAYS_JP = ["月", "火", "水", "木", "金", "土", "日"]

    def shift_all_years_in_text(text):
        return re.sub(r'20(2\d|3[0-5])', lambda m: str(int(m.group(0)) + 1), text)

    def format_weekday(old_str, new_wd_index):
        new_wd = WEEKDAYS_JP[new_wd_index]
        if re.fullmatch(r'[\(（][月火水木金土日][\)）]', old_str.strip()):
            return re.sub(r'[月火水木金土日]', new_wd, old_str)
        if re.fullmatch(r'[月火水木金土日]曜日', old_str.strip()):
            return f"{new_wd}曜日"
        if old_str.strip() in WEEKDAYS_JP:
            return new_wd
        return None

    for sheet in wb.worksheets:
        if year_mode.startswith("すべての年号"):
            new_title = shift_all_years_in_text(sheet.title)
            if new_title != sheet.title:
                sheet.title = new_title
        elif year_mode == "特定の年だけを指定して書き換える" and target_old_year and target_old_year in sheet.title:
            sheet.title = sheet.title.replace(target_old_year, target_new_year)

        sheet_month = None
        m_title = re.search(r'(\d{1,2})月', sheet.title)
        if m_title:
            sheet_month = int(m_title.group(1))

        for row in sheet.iter_rows():
            for cell in row:
                val = cell.value
                if val is None:
                    continue
                if isinstance(val, str) and val.startswith("="):
                    continue

                if isinstance(val, (datetime.datetime, datetime.date)):
                    if year_mode.startswith("すべての年号"):
                        try:
                            cell.value = val.replace(year=val.year + 1)
                        except ValueError:
                            cell.value = val.replace(year=val.year + 1, day=28)
                        updated_date_count += 1
                    elif year_mode == "特定の年だけを指定して書き換える" and target_old_year.isdigit() and target_new_year.isdigit():
                        if val.year == int(target_old_year):
                            try:
                                cell.value = val.replace(year=int(target_new_year))
                            except ValueError:
                                cell.value = val.replace(year=int(target_new_year), day=28)
                            updated_date_count += 1
                    continue

                if isinstance(val, datetime.time) or isinstance(val, bool):
                    continue

                if isinstance(val, str):
                    if year_mode.startswith("すべての年号"):
                        new_val = shift_all_years_in_text(val)
                        if new_val != val:
                            cell.value = new_val
                            updated_date_count += 1
                    elif year_mode == "特定の年だけを指定して書き換える" and target_old_year and target_old_year in val:
                        cell.value = val.replace(target_old_year, target_new_year)
                        updated_date_count += 1
                    continue

                if isinstance(val, (int, float)):
                    if isinstance(val, int) and 2020 <= val <= 2035:
                        if year_mode.startswith("すべての年号"):
                            cell.value = val + 1
                            updated_date_count += 1
                        elif year_mode == "特定の年だけを指定して書き換える" and target_old_year.isdigit() and val == int(target_old_year):
                            cell.value = int(target_new_year)
                            updated_date_count += 1
                        continue

                    is_calendar_day = False
                    if isinstance(val, int) and 1 <= val <= 31:
                        neighbors = [
                            sheet.cell(row=cell.row, column=cell.column + 1),
                            sheet.cell(row=cell.row + 1, column=cell.column)
                        ]
                        for nb in neighbors:
                            if isinstance(nb.value, str) and re.search(r'[月火水木金土日]', nb.value.strip()) and len(nb.value.strip()) <= 4:
                                is_calendar_day = True
                                if not year_mode.startswith("年号・日付は変更しない") and sheet_month and 1 <= sheet_month <= 12:
                                    target_y = int(target_new_year) if (year_mode == "特定の年だけを指定して書き換える" and target_new_year.isdigit()) else 2027
                                    max_day = calendar.monthrange(target_y, sheet_month)
                                    if val <= max_day:
                                        new_wd_idx = datetime.date(target_y, sheet_month, val).weekday()
                                        formatted = format_weekday(nb.value, new_wd_idx)
                                        if formatted:
                                            nb.value = formatted
                                            updated_date_count += 1
                    if is_calendar_day:
                        continue

                    if keep_decimals and isinstance(val, float) and not val.is_integer():
                        continue

                    if cell.column > 1 and keyword_list:
                        left_val = sheet.cell(row=cell.row, column=cell.column - 1).value
                        if isinstance(left_val, str) and any(kw in left_val for kw in keyword_list):
                            continue

                    cell.value = None
                    cleared_count += 1

    default_new_name = shift_all_years_in_text(original_name) if year_mode.startswith("すべての年号") else f"{original_name}_コピー"
    if default_new_name == original_name:
        default_new_name = f"{original_name}_コピー"

    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    b64_data = base64.b64encode(output_buffer.getvalue()).decode()

    st.divider()
    st.success(f"✅ プレビュー準備完了：数値 **{cleared_count} 箇所** をクリアし、年号・日付・曜日を **{updated_date_count} 箇所** 更新しました。")

    html_lines = [
        '<html><head><style>',
        'body { margin: 0; padding: 4px 2px; font-family: sans-serif; color: #31333F; }',
        '.label { font-size: 14px; margin-bottom: 8px; display: block; }',
        '.input-group { display: flex; align-items: center; margin-bottom: 20px; }',
        '.text-input { flex: 1; padding: 10px 12px; font-size: 16px; border: 1px solid #D3D3D3; border-radius: 8px; background-color: #F0F2F6; color: #31333F; outline: none; }',
        '.ext-badge { margin-left: 8px; font-size: 16px; font-weight: bold; color: #555; }',
        '.btn { background-color: #FF4B4B; color: white; border: none; padding: 12px 24px; font-size: 16px; font-weight: bold; border-radius: 8px; cursor: pointer; }',
        '.btn:hover { background-color: #FF3333; }',
        '</style></head><body>',
        '<label class="label">新しいファイル名</label>',
        '<div class="input-group">',
        f'<input type="text" id="filenameInput" class="text-input" value="{default_new_name}">',
        f'<span class="ext-badge">{ext}</span>',
        '</div>',
