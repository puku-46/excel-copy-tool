
import streamlit as st
import io
import os

st.set_page_config(page_title="エクセルコピーツール", layout="centered")

st.title("📄 エクセル完全コピーツール")
st.write("元のエクセルをドラッグ＆ドロップすると、書式・数式・複数タブをそのまま維持してコピーを作成します。")

uploaded_file = st.file_uploader(
    "コピー元のエクセルファイルを選択（またはここにドロップ）", 
    type=["xlsx", "xlsm", "xls"]
)

if uploaded_file is not None:
    original_name, ext = os.path.splitext(uploaded_file.name)
    
    st.divider()
    st.subheader("⚙️ コピー設定")
    
    new_filename = st.text_input(
        "新しいファイル名（拡張子は自動で付きます）", 
        value=f"{original_name}_コピー"
    )
    
    file_bytes = uploaded_file.getvalue()
    copied_buffer = io.BytesIO(file_bytes)
    
    st.download_button(
        label="📥 コピーを作成してダウンロード",
        data=copied_buffer,
        file_name=f"{new_filename}{ext}",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
