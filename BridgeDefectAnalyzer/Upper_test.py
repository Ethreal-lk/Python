
import pandas as pd
from core.defect_rules import DECK_COMPONENT_DEFECT_RULES
from core.excel_handler import read_excel
from core.inspection_filler import process_new_inspection
from core.eport_generator import process_bridge_summary_to_new_file
import os




if __name__ == "__main__":
    # 1. 核心路径配置
    raw_defect_excel_path = r"E:\Python\BridgeDefectAnalyzer\input\病害列表导出2026-10-09.xlsx"
    rule_file_path = r"E:\Python\BridgeDefectAnalyzer\input\城市桥梁评分计算.xlsx"
    output_dir = r"E:\Python\BridgeDefectAnalyzer\output"

    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)

    print(f"1. 正在读取全量病害总表: {raw_defect_excel_path}")
    df_all = read_excel(raw_defect_excel_path, skip_hidden_rows=True)
    print(f"  ✅ 成功读取病害总表，记录数: {len(df_all)} 条")
    df_all.columns = [str(c).strip() for c in df_all.columns]
    print(f"  📋 当前列头为: {df_all.columns.tolist()}")
    # print("部位类型:", df_all["部件类型"].dropna().unique().tolist())
    # print("病害类型:", df_all["病害类型"].dropna().unique().tolist())
        # 检查是否存在“桥梁名称”列
    bridge_col = "桥梁名称"
    if bridge_col not in df_all.columns:
        raise KeyError(f"❌ 原始数据中未找到 '{bridge_col}' 列，无法进行按桥批量处理！当前列头为：{df_all.columns.tolist()}")

    # 2. 按“桥梁名称”进行分组（groupby），实现多桥批量循环
    grouped = df_all.groupby(bridge_col)
    print(f"📊 检测到共有 {len(grouped)} 座桥梁需要处理。\n" + "="*50)
    for bridge_name, df_bridge in grouped:
        # 清理桥梁名称中的非法字符（防止名字带斜杠等导致路径报错）
        safe_bridge_name = str(bridge_name).strip().replace("/", "_").replace("\\", "_")
        print(f"\n🌉 正在处理桥梁: 【 {safe_bridge_name} 】 (包含病害记录: {len(df_bridge)} 条)")
                    # 3. 针对当前桥梁提取最高扣分字典（这里以桥面系为例）
        # deductions_data = parse_structure_defect(df_bridge, "桥面系", defect_map=DEFECT_MAP)