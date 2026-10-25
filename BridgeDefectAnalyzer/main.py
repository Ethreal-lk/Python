import pandas as pd
from core.defect_rules import DECK_COMPONENT_DEFECT_RULES, UPPER_COMPONENT_DEFECT_RULES, LOWER_COMPONENT_DEFECT_RULES,LOWER_COMPONENT_ALIAS
from core.excel_handler import read_excel
from core.inspection_filler import process_new_inspection
from core.eport_generator import process_bridge_summary_to_new_file
from LOwer_test import process_lower_sheet
from QMXTEST2 import process_deck_sheet
from Upper_test import process_upper_sheet
from openpyxl import Workbook,load_workbook
import os
import pandas as pd



if __name__ == "__main__":
    raw_defect_excel_path = r"E:\Python\BridgeDefectAnalyzer\input\病害列表导出2026-10-09_modify1024.xlsx"
    rule_file_path = r"E:\Python\BridgeDefectAnalyzer\input\城市桥梁评分计算.xlsx"
    output_dir = r"E:\Python\BridgeDefectAnalyzer\output"

    os.makedirs(output_dir, exist_ok=True)

    print(f"1. 正在读取全量病害总表: {raw_defect_excel_path}")
    df_all = read_excel(raw_defect_excel_path, skip_hidden_rows=True, skip_hidden_cols=True)
    df_all.columns = [str(c).strip() for c in df_all.columns]

    bridge_col = "桥梁名称"
    if bridge_col not in df_all.columns:
        raise KeyError(f"❌ 原始数据中未找到 '{bridge_col}' 列，当前列头为：{df_all.columns.tolist()}")

    grouped = df_all.groupby(bridge_col)
    print(f"📊 检测到共有 {len(grouped)} 座桥梁需要处理。\n" + "=" * 50)

    for bridge_name, df_bridge in grouped:
        safe_bridge_name = str(bridge_name).strip().replace("/", "_").replace("\\", "_")
        print(f"\n🌉 正在处理桥梁: 【 {safe_bridge_name} 】 (病害记录: {len(df_bridge)} 条)")

        wb = Workbook()

        # Sheet 1: 桥面系
        process_deck_sheet(wb, df_bridge, rule_file_path, safe_bridge_name)

        # Sheet 3: 下部结构
        process_lower_sheet(wb, df_bridge, rule_file_path, safe_bridge_name)

        # Sheet 2: 上部结构
        process_upper_sheet(wb, df_bridge, rule_file_path, safe_bridge_name)

        out_path = os.path.join(output_dir, f"{safe_bridge_name}_评分表.xlsx")
        wb.save(out_path)
        print(f"  📁 已保存: {out_path}")

    print("\n" + "=" * 50)
    print("🎉 所有桥梁处理完毕！")