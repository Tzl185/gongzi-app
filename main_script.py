from openpyxl import load_workbook
import os
import pandas as pd

def get_flat_column_names(filepath, header_row=4):
    """
    自动读取合并单元格表头，生成扁平化列名。
    默认第4行为表头（从1开始计数）
    """
    wb = load_workbook(filepath, data_only=True)
    ws = wb.active

    merged_dict = {}
    for merged_range in ws.merged_cells.ranges:
        min_col = merged_range.min_col
        min_row = merged_range.min_row
        max_col = merged_range.max_col
        max_row = merged_range.max_row

        merged_value = ws.cell(row=min_row, column=min_col).value
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                merged_dict[(row, col)] = merged_value

    col_names = []
    for col in range(1, ws.max_column + 1):
        upper = merged_dict.get((header_row - 1, col)) or ws.cell(row=header_row - 1, column=col).value
        lower = merged_dict.get((header_row, col)) or ws.cell(row=header_row, column=col).value

        parts = []
        if upper: parts.append(str(upper).strip())
        if lower and lower != upper: parts.append(str(lower).strip())
        col_name = "-".join(parts) if parts else f"列{col}"
        col_names.append(col_name)

    return col_names

def process_file_a(folder_path, output_file="文件A_汇总结果.xlsx"):
    """
    自动处理文件夹中的Excel，支持合并单元格的表头，生成文件A并返回数据日志
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

                # 获取扁平列名
                columns = get_flat_column_names(filepath, header_row=4)
                df = pd.read_excel(filepath, header=3)
                df.columns = columns
                df = df.reset_index(drop=True)

                print(f"{filename} 的列名如下:", file=log)
                print(columns, file=log)

                # 假设预算单位是第2列
                budget_unit_col = columns[1]
                wage_cols = columns[16:30]  # Q~AD列

                df_filtered = df[[budget_unit_col] + wage_cols]
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

                if not df_grouped.empty:
                    all_data.append(df_grouped)

            except Exception as e:
                print(f"处理文件 {filename} 出错: {e}", file=log)

    if all_data:
        df_all = pd.concat(all_data)
        df_final = df_all.groupby(df_all.index).sum()
        output_path = os.path.join(folder_path, output_file)
        df_final.to_excel(output_path)
        print(f"\n汇总结果已保存到: {output_path}", file=log)
        print(f"共收集到 {len(all_values)} 项工资数据", file=log)
        return output_path, all_values, log.getvalue()
    else:
        print("没有找到有效数据，请检查格式或表头", file=log)
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
                    break

        output_path = os.path.join(os.path.dirname(file_b_path), "updated_" + os.path.basename(file_b_path))
        wb.save(output_path)
        print(f"\n共匹配成功 {match_count} 项工资数据")
        return output_path

    except Exception as e:
        print(f"\n更新文件B出错: {e}")
        return None
