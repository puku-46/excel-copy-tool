import streamlit as st
import streamlit.components.v1 as components
import openpyxl, io, base64, os, re, datetime, calendar

st.set_page_config(page_title="エクセル白紙化コピーツール", layout="centered")
st.title("📄 エクセル白紙化コピーツール")

uploaded_file = st.file_uploader("エクセルを選択（またはドロップ）", type=["xlsx", "xlsm"])

if uploaded_file is not None:
    original_name, ext = os.path.splitext(uploaded_file.name)
    keep_vba = (ext.lower() == ".xlsm")
    wb = openpyxl.load_workbook(uploaded_file, keep_vba=keep_vba)

    # 1. 事前スキャン
    found_years = set()
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and not v.startswith("="):
                    found_years.update(re.findall(r'20\d{2}', v))
                elif isinstance(v, (datetime.datetime, datetime.date)):
                    found_years.add(str(v.year))
                elif isinstance(v, int) and 2020 <= v <= 2035:
                    found_years.add(str(v))

    st.divider()
    st.subheader("🔍 1. 事前チェック：検出された年号")
    st.info(f"検出された年号: **{', '.join(sorted(found_years)) if found_years else 'なし'}**")

    year_mode = st.radio("📅 年号・日付の変更方式：", [
        "すべての年号・日付を ＋1年 スライド（例: 2025→2026, 2026→2027, 2027→2028）",
        "特定の年だけ指定して書き換える",
        "変更しない"
    ])

    old_y, new_y = "2026", "2027"
    if year_mode == "特定の年だけ指定して書き換える":
        c1, c2 = st.columns(2)
        old_y = c1.text_input("変更前", "2026")
        new_y = c2.text_input("変更後", "2027")

    st.divider()
    st.subheader("🧹 2. 数値クリアのルール")
    c3, c4 = st.columns(2)
    keep_dec = c3.checkbox("小数の数値（単価など）は残す", value=True)
    kw_str = c4.text_input("残したい項目のキーワード", value="基本料金,単価,単年割,割合,燃調費,再エネ,市場価格")
    kw_list = [k.strip() for k in kw_str.split(",") if k.strip()]

    # 2. 変換処理
    cleared, updated = 0, 0
    shift_year = lambda s: re.sub(r'20(2\d|3[0-5])', lambda m: str(int(m.group(0)) + 1), s)

    for sheet in wb.worksheets:
        if year_mode.startswith("すべての年号"):
            sheet.title = shift_year(sheet.title)
        elif year_mode == "特定の年だけ指定して書き換える" and old_y:
            sheet.title = sheet.title.replace(old_y, new_y)

        for row in sheet.iter_rows():
            for cell in row:
                v = cell.value
                if v is None or (isinstance(v, str) and v.startswith("=")) or isinstance(v, (datetime.time, bool)):
                    continue
                if isinstance(v, (datetime.datetime, datetime.date)):
                    if year_mode.startswith("すべての年号"):
                        cell.value = v.replace(year=v.year + 1, day=min(v.day, 28) if v.month == 2 else v.day)
                        updated += 1
                    elif year_mode == "特定の年だけ指定して書き換える" and old_y.isdigit() and v.year == int(old_y):
                        cell.value = v.replace(year=int(new_y), day=min(v.day, 28) if v.month == 2 else v.day)
                        updated += 1
                elif isinstance(v, str):
                    nv = shift_year(v) if year_mode.startswith("すべての年号") else (v.replace(old_y, new_y) if year_mode == "特定の年だけ指定して書き換える" and old_y else v)
                    if nv != v:
                        cell.value = nv
                        updated += 1
                elif isinstance(v, (int, float)):
                    if isinstance(v, int) and 2020 <= v <= 2035:
                        if year_mode.startswith("すべての年号"):
                            cell.value, updated = v + 1, updated + 1
                        elif year_mode == "特定の年だけ指定して書き換える" and old_y.isdigit() and v == int(old_y):
                            cell.value, updated = int(new_y), updated + 1
                        continue
                    if keep_dec and isinstance(v, float) and not v.is_integer():
                        continue
                    # 左隣または2つ左のセルに保護キーワードがあれば残す（セル結合対策）
                    protected = False
                    for offset in:
                        if cell.column > offset:
                            left_v = str(sheet.cell(cell.row, cell.column - offset).value or "")
                            if any(k in left_v for k in kw_list):
                                protected = True
                    if protected:
                        continue
                    cell.value = None
                    cleared += 1

    def_name = shift_year(original_name) if year_mode.startswith("すべての年号") else f"{original_name}_コピー"
    if def_name == original_name: def_name += "_コピー"

    buf = io.BytesIO()
    wb.save(buf)
    b64 = base64.b64encode(buf.getvalue()).decode()

    st.success(f"✅ 完了：数値 **{cleared} 箇所** クリア / 年号・日付 **{updated} 箇所** 更新")
    html = f'<input id="fn" value="{def_name}" style="padding:8px;width:65%;font-size:16px;"> <b>{ext}</b><br><br><button onclick="dl()" style="background:#FF4B4B;color:#fff;border:none;padding:10px 20px;font-size:16px;border-radius:6px;cursor:pointer;">📥 ダウンロード</button><script>function dl(){{let n=document.getElementById("fn").value||"{def_name}";let b=atob("{b64}"),a=new Uint8Array(b.length);for(let i=0;i<b.length;i++)a[i]=b.charCodeAt(i);let u=URL.createObjectURL(new Blob([a],{{type:"application/octet-stream"}}));let l=document.createElement("a");l.href=u;l.download=n+"{ext}";l.click();}}</script>'
    components.html(html, height=110)
