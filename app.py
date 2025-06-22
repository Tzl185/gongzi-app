# app.py
import streamlit as st
import zipfile, tempfile, os
import pandas as pd
from main_script import process_file_a, update_file_b, convert_xls_to_xlsx

st.set_page_config(page_title="工资自动处理", layout="wide")
st.title("📊 工资数据自动汇总与更新系统")
log_placeholder = st.empty()

st.markdown("""
1. 上传 `.zip` 文件，里面放多个工资 `.xls` 或 `.xlsx` 表格（标题在第4行）
2. 自动生成“文件A”汇总结果
3. 可选：上传模板文件B，用文件A数值自动更新J列
""")

zip_file = st.file_uploader("📂 上传包含多个Excel的压缩包 (.zip)", type="zip")
file_b = st.file_uploader("📄 （可选）上传模板文件B (.xlsx)", type="xlsx")

if st.button("🚀 开始处理"):
    if not zip_file:
        st.error("请先上传zip文件")
    else:
        with tempfile.TemporaryDirectory() as tmpdir:
            zip_path = os.path.join(tmpdir, "uploaded.zip")
            with open(zip_path, "wb") as f:
                f.write(zip_file.read())

            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(tmpdir)

            # 找到解压后的子目录（用户上传的内容）
            folder_a = None
            for root, dirs, files in os.walk(tmpdir):
                for f in files:
                    if f.endswith(('.xls', '.xlsx')):
                        folder_a = root
                        break
                if folder_a:
                    break

            if not folder_a:
                st.error("压缩包中没有找到Excel文件")
            else:
                # 自动转xls为xlsx
                for fname in os.listdir(folder_a):
                    if fname.endswith('.xls') and not fname.startswith('~$'):
                        old_path = os.path.join(folder_a, fname)
                        new_path = convert_xls_to_xlsx(old_path)
                        os.remove(old_path)

                st.write(f"📁 解压目录：{folder_a}")
                log = []
                def logger(msg):
                    log.append(msg)
                    log_placeholder.code("\n".join(log[-20:]), language="text")

                logger("📥 正在分析并生成文件A...")
                try:
                    file_a_path, values = process_file_a(folder_a, logger=logger)
                    if not file_a_path:
                        st.error("❌ 文件A处理失败，请检查Excel格式或第4行列名")
                    else:
                        with open(file_a_path, "rb") as f:
                            st.download_button("⬇️ 下载文件A（汇总结果）", f, file_name="文件A_汇总结果.xlsx")

                        # 第二步更新文件B
                        if file_b:
                            logger("\n📤 正在用文件A更新文件B...")
                            b_path = os.path.join(tmpdir, "uploaded_b.xlsx")
                            with open(b_path, "wb") as f:
                                f.write(file_b.read())
                            updated_path = update_file_b(file_a_path, b_path, logger=logger)
                            if updated_path:
                                with open(updated_path, "rb") as f:
                                    st.download_button("⬇️ 下载更新后的文件B", f, file_name="updated_模板.xlsx")
                except Exception as e:
                    st.error(f"处理时出错: {e}")
