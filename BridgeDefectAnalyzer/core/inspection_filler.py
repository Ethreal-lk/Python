import re
from openpyxl import load_workbook


def parse_excel_if_formula(formula_str):
    """
    解析 Excel 中的嵌套 IF 公式：
    1. 提取扣分值 -> 损坏程度文本的映射表 (score_to_degree)
    2. 提取 扣分=0 时对应的默认初始化文本 (default_degree，如 '完好'、'正常'、'足够'、'无')
    """
    if not isinstance(formula_str, str) or not formula_str.startswith("="):
        return {}, "无"

    # 匹配类似于 Cxx="文本",数值 的正则模式
    pattern = r'C\d+\s*=\s*"([^"]+)"\s*,\s*(\d+)'
    matches = re.findall(pattern, formula_str)

    score_to_degree = {}
    default_degree = "无"  # 兜底默认值

    for degree_text, score in matches:
        score_val = int(score)
        score_to_degree[score_val] = degree_text
        if score_val == 0:
            default_degree = degree_text  # 动态捕捉 0 扣分对应的真正默认词！

    return score_to_degree, default_degree


def process_new_inspection(template_path, target_deductions, sheet_name="桥面系评分"):
    print("1. 正在解析 Excel 公式、扣分规则及 0 扣分默认初始值...")
    wb_read = load_workbook(template_path, data_only=False)
    ws_read = wb_read[sheet_name]

    defect_rules = {}      # 病害扣分映射规则
    default_degrees = {}   # 记录每种病害 0 扣分时的默认文本（如 '完好'/'正常'/'足够'/'无'）
    row_mapping = {}       # 行号映射

    for row_idx, row in enumerate(ws_read.iter_rows(min_row=10, values_only=False), start=10):
        defect_type = row[1].value  # B列：损坏类型
        formula = row[3].value      # D列：单项扣分公式

        if defect_type and formula:
            rule_map, default_deg = parse_excel_if_formula(str(formula))
            if rule_map:
                defect_rules[defect_type] = rule_map
                default_degrees[defect_type] = default_deg
                row_mapping[defect_type] = row_idx

    wb_read.close()
    print(f"   └─ 成功加载 {len(defect_rules)} 条标准规则。")

    print("\n2. 执行【全表精准动态初始化】：利用 0 扣分反填全表 C 列...")
    wb_write = load_workbook(template_path)
    ws_write = wb_write[sheet_name]

    # 根据每项病害公式里的 0 分对应值，精准重置 C 列
    for defect_type, row_num in row_mapping.items():
        reset_text = default_degrees[defect_type]
        ws_write.cell(row=row_num, column=3, value=reset_text)
        
    print("   └─ 精准初始化完成！防水层已置为'完好'，防滑能已置为'足够'，缝宽已置为'正常'，其余已置为'无'。")

    print("\n3. 写入本次全新的检测病害数据...")
    updated_count = 0

    for defect_type, target_score in target_deductions.items():
        if defect_type not in defect_rules:
            print(f"⚠️  [跳过] 未在模板中找到病害【{defect_type}】的规则")
            continue

        rule_map = defect_rules[defect_type]
        row_num = row_mapping[defect_type]

        # 匹配扣分对应的损坏程度文本
        if target_score in rule_map:
            degree_text = rule_map[target_score]
        else:
            closest_score = min(rule_map.keys(), key=lambda k: abs(k - target_score))
            degree_text = rule_map[closest_score]
            print(f"💡 [自动纠偏] 病害【{defect_type}】目标扣分 {target_score} 无精准匹配 -> 修正为规则分值 {closest_score} ('{degree_text}')")

        # 写入 C 列（损坏程度）
        ws_write.cell(row=row_num, column=3, value=degree_text)
        updated_count += 1
        print(f"✅ [填入成功] 行号 {row_num:<2} | 病害: {defect_type:<14} | 扣分: {target_score:<2} -> 损坏程度: '{degree_text}'")

    # 另存为副本文件
    output_path = template_path.replace(".xlsx", "_计算结果.xlsx")
    wb_write.save(output_path)
    wb_write.close()

    print(f"\n🎉 场景 A 全新检测填报完成！更新了 {updated_count} 项有扣分的病害。")
    print(f"📁 结果文件已保存至: {output_path}")
    # 🚀 将生成的计算结果文件路径返回，供后续接口或主流程使用
    return output_path 