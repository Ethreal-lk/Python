import pandas as pd
from core.defect_rules import DECK_COMPONENT_DEFECT_RULES
from core.excel_handler import read_excel
from core.inspection_filler import process_new_inspection
from core.eport_generator import process_bridge_summary_to_new_file

# 构建 (部件名称, 现场病害) -> 规范大类 的反向查找字典
DEFECT_MAP = {
    (component, sub_defect): std_category
    for component, categories in DECK_COMPONENT_DEFECT_RULES.items()
    for std_category, sub_defects in categories.items()
    for sub_defect in sub_defects
}

UPPER_DEFECT_MAP = {

}

LOWER_DEFECT_MAP = {

}

def parse_structure_defect(df_all, target_part, defect_map):
    """
    通用结构部位病害解析与大类扣分计算器
    :param df_all: 清洗后的总表 DataFrame
    :param target_part: 构件部位（如 "桥面系", "上部结构", "下部结构"）
    :param defect_map: 对应部位的反向映射字典
    """
    print(f"  └─ 正在处理子模块: 【{target_part}】...")
    
    # 1. 过滤指定部位且扣分 > 0 的记录
    part_df = df_all[(df_all["部位类型"] == target_part) & (df_all["扣分分数"] > 0)].copy()
    print(f"     [ {target_part} ] 筛选出有效扣分记录 {len(part_df)} 条。")

    target_deductions = {}

    # 2. 遍历并映射取最大扣分
    for _, row in part_df.iterrows():
        component = str(row["部件类型"]).strip() if pd.notna(row["部件类型"]) else ""
        raw_defect = str(row["病害类型"]).strip() if pd.notna(row["病害类型"]) else ""
        score = int(row["扣分分数"])

        std_category = defect_map.get((component, raw_defect))

        if std_category:
            current_max = target_deductions.get(std_category, 0)
            target_deductions[std_category] = max(current_max, score)
        else:
            target_deductions[raw_defect] = max(target_deductions.get(raw_defect, 0), score)

    return target_deductions

# if __name__ == "__main__":
#     raw_defect_excel_path = r"G:\Python_lk\Python\BridgeDefectAnalyzer\input\病害列表导出2026-10-09.xlsx"
#     rule_file_path = r"G:\Python_lk\Python\BridgeDefectAnalyzer\input\城市桥梁评分计算.xlsx"
#     output_path = (
#       r"G:\Python_lk\Python\BridgeDefectAnalyzer\output\桥面系各要素扣分汇总.xlsx"
#     )
#     #     print(f"1. 正在读取原始病害明细表: {raw_excel_path}")
#     df = read_excel(raw_defect_excel_path,skip_hidden_rows=True)
#     print(df)

#     # 提取最高扣分字典
#     deductions_data = parse_structure_defect(df,"桥面系",defect_map = DEFECT_MAP)
#     print(deductions_data)
    
#     # 场景 A 初始化复位并反填
#     Pingfen_path= process_new_inspection(template_path = rule_file_path, target_deductions=deductions_data)
#     print(Pingfen_path)

#     process_bridge_summary_to_new_file(Pingfen_path, output_path)

import os
import pandas as pd

if __name__ == "__main__":
    # 1. 核心路径配置
    raw_defect_excel_path = r"G:\Python_lk\Python\BridgeDefectAnalyzer\input\病害列表导出2026-10-09.xlsx"
    rule_file_path = r"G:\Python_lk\Python\BridgeDefectAnalyzer\input\城市桥梁评分计算.xlsx"
    output_dir = r"G:\Python_lk\Python\BridgeDefectAnalyzer\output"

    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)

    print(f"1. 正在读取全量病害总表: {raw_defect_excel_path}")
    df_all = read_excel(raw_defect_excel_path, skip_hidden_rows=True)
    df_all.columns = [str(c).strip() for c in df_all.columns]

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

        try:
            # 3. 针对当前桥梁提取最高扣分字典（这里以桥面系为例）
            deductions_data = parse_structure_defect(df_bridge, "桥面系", defect_map=DEFECT_MAP)
            
            if not deductions_data:
                print(f"  ⚠️ 桥梁 [{safe_bridge_name}] 没有筛选到有效的桥面系扣分记录，跳过生成。")
                continue

            # 4. 场景 A 初始化复位并反填（传入模板和当前桥的数据）
            pingfen_path = process_new_inspection(template_path=rule_file_path, target_deductions=deductions_data)

            # 5. 动态生成以桥名命名的输出文件路径
            bridge_output_path = os.path.join(output_dir, f"{safe_bridge_name}_桥面系扣分汇总.xlsx")

            # 6. 生成最终汇总文件
            process_bridge_summary_to_new_file(pingfen_path, bridge_output_path)
            print(f"  ✅ 桥梁 [{safe_bridge_name}] 处理成功！文件已保存至: {bridge_output_path}")

        except Exception as e:
            print(f"  ❌ 桥梁 [{safe_bridge_name}] 处理失败，错误信息: {e}")
            continue

    print("\n" + "="*50)
    print("🎉 所有桥梁的批量自动化评估流水线已全部执行完毕！")