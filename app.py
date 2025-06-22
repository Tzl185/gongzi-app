import streamlit as st
import zipfile
import os
import tempfile
import shutil
from main_script import process_file_a, update_file_b

st.set_page_config(page_title="工资调整工具", layout="wide")
st.title("📊 工资调整自动处理工具")

uploaded_zip = st.file_uploader("上传包含多个 Excel 的 zip 文件（文件夹A）", type=["zip"])
uploaded_file_b = st.file_uploader("上传模板文件B（项目细化导入模板）", type=["xlsx"])

if st.button("开始处理") and uploaded_zip and uploaded_file_b:
    with tempfile.TemporaryDirectory() as tmpdir:
       # 解压zip为文件夹A
        zip_path = os.path.join(tmpdir, "folder_a.zip")
        with open(zip_path, "wb") as f:
            f.write(uploaded_zip.read())
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            folder_a_path = os.path.join(tmpdir, "folder_a")
            zip_ref.extractall(folder_a_path)
           
        st.write("🗂️ 解压后的文件列表：")
        for root, dirs, files in os.walk(folder_a_path):
            for name in files:
                st.write(os.path.join(root, name))
        
        # 保存模板文件B
        file_b_path = os.path.join(tmpdir, "file_b.xlsx")
        with open(file_b_path, "wb") as f:
            f.write(uploaded_file_b.read())

        # 处理文件A
        st.write("📥 正在处理文件A...")
        file_a_path, all_values = process_file_a(folder_a_path)

        if file_a_path and all_values:
            st.success("✅ 文件A生成成功")

            # 更新文件B
            st.write("🔄 正在更新文件B...")
            updated_b_path = update_file_b(file_a_path, file_b_path)

            if updated_b_path:
                with open(updated_b_path, "rb") as f:
                    st.download_button("📥 下载更新后的文件B", f, file_name="updated_文件B.xlsx")
            else:
                st.error("❌ 文件B更新失败")
        else:
            st.error("❌ 文件A处理失败，请检查Excel格式")
