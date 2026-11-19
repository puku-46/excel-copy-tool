import streamlit as st
import os

# ページの基本設定
st.set_page_config(page_title="エクセルコピーツール", layout="centered")

st.title("📄 エクセル完全コピーツール")
st.write("元のエクセルをドラッグ＆ドロップすると、書式・数式・複数タブをそのまま維持してコピーを作成します。")

# 1. ファイルのアップロード（ドラッグ＆ドロップ対応）
uploaded_file = st.file_uploader(
    "コピー元のエクセルファイルを選択（またはここにドロップ）", 
    type=["xlsx", "xlsm", "xls"]
)

if uploaded_file is not None:
    # 元のファイル名と拡張子を取得
    original_name, ext = os.path.splitext(uploaded_file.name)
    
    st.divider()
    st.subheader("⚙️ コピー設定")
    
    # 2. 新しいファイル名の指定
    new_filename = st.text_input(
        "新しいファイル名（拡張子は自動で付きます）", 
        value=f"{original_name}_コピー"
    )
    
    # 3. データをバイト列（bytes）としてそのまま取得（何回ダウンロードしても消えない形式）
    file_bytes = uploaded_file.getvalue()
    
    # 4. ダウンロードボタンの表示
    st.download_button(
        label="📥 コピーを作成してダウンロード",
        data=file_bytes,
        file_name=f"{new_filename}{ext}",
        mime="application/octet-stream",
        type="primary"
    )
