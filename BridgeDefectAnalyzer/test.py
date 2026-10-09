import pandas as pd
from core.defect_rules import DECK_COMPONENT_DEFECT_RULES
from core.excel_handler import read_excel
from rule_test import process_new_inspection

# 构建 (部件名称, 现场病害) -> 规范大类 的反向查找字典
DEFECT_MAP = {
    (component, sub_defect): std_category
    for component, categories in DECK_COMPONENT_DEFECT_RULES.items()
    for std_category, sub_defects in categories.items()
    for sub_defect in sub_defects
}

def parse_raw_defect_excel(raw_excel_path):
    """
    读取现场原始病害 Excel 明细表，清洗非数值扣分并完成归类反推
    """
    print(f"1. 正在读取原始病害明细表: {raw_excel_path}")
    df = read_excel(raw_excel_path,skip_hidden_rows=True)
    
    # 清理列名两端可能存在的空格
    df.columns = [str(c).strip() for c in df.columns]

    # 根据控制台输出的真实列名设置
    part_col = "部位类型"
    component_col = "部件类型"
    defect_col = "病害类型"
    score_col = "扣分分数"

    # ---------------- 关键修复：清洗扣分列数据类型 ----------------
    # 将扣分列强制转为数值类型，遇到非数字（如空值、文本）自动转为 0
    df[score_col] = pd.to_numeric(df[score_col], errors='coerce').fillna(0)

    # 1.1 过滤：仅保留桥面系且扣分 > 0 的有效病害记录
    qmx_df = df[(df[part_col] == "桥面系") & (df[score_col] > 0)].copy()
    print(f"   └─ 成功筛选出桥面系有效扣分记录 {len(qmx_df)} 条。")

    target_deductions = {}

    # 1.2 遍历明细，完成映射并取同大类最大扣分
    for _, row in qmx_df.iterrows():
        component = str(row[component_col]).strip() if pd.notna(row[component_col]) else ""
        raw_defect = str(row[defect_col]).strip() if pd.notna(row[defect_col]) else ""
        score = int(row[score_col])

        # 从映射字典中查找规范大类名称
        std_category = DEFECT_MAP.get((component, raw_defect))

        if std_category:
            # 同一大类存在多项扣分时取最高分
            current_max = target_deductions.get(std_category, 0)
            target_deductions[std_category] = max(current_max, score)
            print(f"  ├─ [精准映射] [{component}] '{raw_defect}' ({score}分) -> 规范大类 '{std_category}' (更新当前最高扣分: {target_deductions[std_category]}分)")
        else:
            # 未精准匹配时保留原名称兜底
            target_deductions[raw_defect] = max(target_deductions.get(raw_defect, 0), score)
            print(f"  ├─ ⚠️ [未精准匹配] [{component}] '{raw_defect}'，暂保留原名称")

    print(f"\n2. 原始数据提取完毕！最终转换得到的扣分字典:\n   {target_deductions}")
    return target_deductions


if __name__ == "__main__":
    raw_defect_excel_path = r"E:\Python\BridgeDefectAnalyzer\input\virus.xlsx"
    rule_file_path = r"E:\Python\BridgeDefectAnalyzer\input\城市桥梁评分计算.xlsx"
    
    # 提取最高扣分字典
    deductions_data = parse_raw_defect_excel(raw_defect_excel_path)
    
    # 场景 A 初始化复位并反填
    process_new_inspection(template_path = rule_file_path, target_deductions=deductions_data)