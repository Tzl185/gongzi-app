from openpyxl import load_workbook
import os
import pandas as pd

def process_file_a(folder_path, output_file="文件A_汇总结果.xlsx"):
    """
    处理文件夹中的Excel表，生成汇总的文件A
    返回文件A的路径和所有数值字典，及日志字符串
    """
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

        output_path = os.path.join(folder_path, output_file)
        df_final.to_excel(output_path)
        print(f"\n汇总结果已保存到: {output_path}", file=log)

        print(f"\n总共收集到 {len(all_values)} 个数值", file=log)
        return output_path, all_values, log.getvalue()
    else:
        print("没有找到有效数据", file=log)
        return None, None, log.getvalue()


def update_file_b(file_a_path, file_b_path):
    """
    用文件A中的所有数值更新文件B的J列，保留原有格式
    """
    try:
        df_a = pd.read_excel(file_a_path, index_col=0)
        wage_cols = df_a.columns

        all_values = {}
        for budget_unit, row in df_a.iterrows():
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

        wb = load_workbook(file_b_path)
        sheet = wb.active

        j_col_index = 10  # J列

        match_count = 0
        for row_idx in range(2, sheet.max_row + 1):
            unit_cell = sheet.cell(row=row_idx, column=1)
            unit_info = str(unit_cell.value).strip() if unit_cell.value else ""

            budget_cell = sheet.cell(row=row_idx, column=2)
            budget_project = str(budget_cell.value).strip() if budget_cell.value else ""

            unit_info_cleaned = unit_info.replace("-", "").replace(" ", "")

            matched = False
            for (budget_unit, wage_type), value in all_values.items():
                budget_unit_cleaned = budget_unit.replace("-", "").replace(" ", "")

                unit_match = (budget_unit_cleaned in unit_info_cleaned) or (unit_info_cleaned in budget_unit_cleaned)
                wage_match = wage_type in budget_project

                if unit_match and wage_match:
                    sheet.cell(row=row_idx, column=j_col_index).value = value
                    match_count += 1
                    matched = True
                    print(f"匹配成功: 行{row_idx} 单位:'{unit_info}'⊇'{budget_unit}', 项目:'{budget_project}'⊇'{wage_type}', 值:{value}")
                    break

            if not matched and row_idx < 7:
                print(f"未匹配: 行{row_idx} 单位:'{unit_info}', 项目:'{budget_project}'")

        output_path = os.path.join(os.path.dirname(file_b_path), "updated_" + os.path.basename(file_b_path))
        wb.save(output_path)
        print(f"\n总共完成 {match_count} 处匹配")
        print(f"已保存更新后的文件B到: {output_path}")
        return output_path

    except Exception as e:
        print(f"\n更新文件B出错: {e}")
        return None
