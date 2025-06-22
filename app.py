import streamlit as st
import tempfile
import zipfile
import os
import pandas as pd
from main_script import process_file_a, update_file_b

st.title("工资预算处理工具")

uploaded_zip = st.file_uploader("上传压缩的文件夹（ZIP格式，内含Excel文件）", type="zip")
uploaded_file_b = st.file_uploader("上传文件B（模板Excel）", type=["xls", "xlsx"])

def run_process_file_a(folder_path):
    import io
    log = io.StringIO()

    all_data = []
    all_values = {}
    
    for filename in os.listdir(folder_path):
        if filename.endswith(('.xls', '.xlsx')) and not filename.startswith('~$'):
            filepath = os.path.join(folder_path, filename)
            try:
                print(f"开始处理文件: {filename}", file=log)

                df_raw = pd.read_excel(filepath, header=None)
                df = df_raw[3:]
                df.columns = df_raw.iloc[3]
                df = df.reset_index(drop=True)

                print(f"{filename} 的列名如下:", file=log)
                print(list(df.columns), file=log)

                budget_unit_col = df.columns[1]
                wage_cols = df.columns[16:30]

                df_filtered = df[[budget_unit_col] + list(wage_cols)]
                df_filtered[wage_cols] = df_filtered[wage_cols].apply(pd.to_numeric, errors='coerce').fillna(0)

                df_grouped = df_filtered.groupby(budget_unit_col).sum()

                for budget_unit, row in df_grouped.iterrows():
                    for wage_type in wage_cols:
                        value = row[wage_type]
                        wage_type_str = str(wage_type).strip()
                        if "绩效工资" in wage_type_str:
                            wage_type_str = wage_type_str.replace("绩效工资", "基础性绩效")
                        if "行政医疗" in wage_type_str:
                            wage_type_str = wage_type_str.replace("行政医疗", "职工基本医疗（行政）")
                        elif "事业医疗" in wage_type_str:
                            wage_type_str = wage_type_str.replace("事业医疗", "基本医疗（事业）")
                        elif "医疗保险" in wage_type_str:
                            wage_type_str = wage_type_str.replace("医疗保险", "基本医疗")

                        key = (str(budget_unit).strip(), wage_type_str)
                        all_values[key] = value
                        if "医疗" in wage_type_str:
                            print(f"医疗数值记录 - 单位: {budget_unit}, 类型: {wage_type_str}, 值: {value}", file=log)

                if df_grouped is not None and not df_grouped.empty:
                    all_data.append(df_grouped)
                
            except Exception as e:
                print(f"处理文件 {filename} 出错: {e}", file=log)
    
    if all_data:
        df_all = pd.concat(all_data)
        df_final = df_all.groupby(df_all.index).sum()
        
        output_path = os.path.join(folder_path, "文件A_汇总结果.xlsx")
        df_final.to_excel(output_path)
        print(f"\n汇总结果已保存到: {output_path}", file=log)
        
        print(f"\n总共收集到 {len(all_values)} 个数值", file=log)
        return output_path, all_values, log.getvalue()
    else:
        print("没有找到有效数据", file=log)
        return None, None, log.getvalue()

if uploaded_zip and uploaded_file_b:
    with tempfile.TemporaryDirectory() as tmpdir:
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

        file_b_path = os.path.join(tmpdir, "file_b.xlsx")
        with open(file_b_path, "wb") as f:
            f.write(uploaded_file_b.read())

        if st.button("开始处理"):
            file_a_path, all_values, log_text = run_process_file_a(folder_a_path)
            st.text_area("📋 文件A处理日志", log_text, height=300)

            if file_a_path and all_values:
                st.success(f"文件A生成成功: {file_a_path}")
                updated_file_b_path = update_file_b(file_a_path, file_b_path)
                if updated_file_b_path:
                    st.success(f"文件B更新成功，保存为：{updated_file_b_path}")
                else:
                    st.error("文件B更新失败，请检查日志。")
            else:
                st.error("❌ 文件A处理失败，请检查Excel格式和日志信息。")
