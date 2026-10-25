import pandas as pd
from core.defect_rules import (
    DECK_COMPONENT_DEFECT_RULES,
    DECK_COMPONENT_ALIAS,
    LOWER_COMPONENT_DEFECT_RULES,
    LOWER_COMPONENT_ALIAS,
    LOWER_DEFECT_MAP,
    DECK_COMPONENT_TYPES,
    UPPER_COMPONENT_TYPES,
    UPPER_COMPONENT_ALIAS,
    UPPER_COMPONENT_DEFECT_RULES

)
from core.excel_handler import read_excel
from core.eport_generator import process_bridge_summary_to_new_file
from core.inspection_filler import parse_excel_if_formula
from openpyxl import load_workbook
import os
import sys
import re
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
from LOwer_test import traverse_bridge_data,parse_template_rules,score_to_degree,fill_missing_defects

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")



UPPER_PART_DEFECT_TYPES = {
    component: list(categories.keys())
    for component, categories in UPPER_COMPONENT_DEFECT_RULES.items()
}
def draw_upper_structure(ws, df_upper):
    current_row = 1

    # 按部件遍历（PC或RC梁式构件、横向联系、主拱圈等）
    for part in df_upper["部件类型"].unique():
        defect_order = UPPER_PART_DEFECT_TYPES.get(part, [])
        if not defect_order:
            continue

        df_part = df_upper[df_upper["部件类型"] == part]
        score_map = dict(zip(df_part["损坏类型"], df_part["扣分分数"]))
        degree_map = dict(zip(df_part["损坏类型"], df_part["程度"]))

        r1 = current_row

        # 表头
        ws.merge_cells(start_row=r1, start_column=1, end_row=r1, end_column=2)
        ws.cell(row=r1, column=1, value="损害类型\n结构位置")
        ws.cell(row=r1, column=1).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        
        for idx, defect in enumerate(defect_order):
            col = 3 + idx
            cell = ws.cell(row=r1, column=col, value=defect)
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # 部件名合并两行
        ws.merge_cells(start_row=r1 + 1, start_column=1, end_row=r1 + 2, end_column=1)
        ws.cell(row=r1 + 1, column=1, value=part)
        ws.cell(row=r1 + 1, column=1).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=r1 + 1, column=2, value="程度")
        ws.cell(row=r1 + 2, column=2, value="扣分值")

        # 数据行
        for idx, defect in enumerate(defect_order):
            col = 3 + idx
            score = score_map.get(defect, 0)
            degree = degree_map.get(defect, "无")
            ws.cell(row=r1 + 1, column=col, value=degree)
            ws.cell(row=r1 + 2, column=col, value=score)

        # 边框与样式
        thin = Side(style="thin")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        max_col = 2 + len(defect_order)

        for r in range(r1, r1 + 3):
            for c in range(1, max_col + 1):
                cell = ws.cell(row=r, column=c)
                cell.border = border
                if not (r == r1 and c == 1):
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.font = Font(name="宋体", size=11, bold=(r == r1 or c <= 2))

        diag = Border(left=thin, right=thin, top=thin, bottom=thin, diagonal=thin, diagonalDown=True)
        ws.cell(row=r1, column=1).border = diag
        ws.cell(row=r1, column=2).border = diag

        ws.row_dimensions[r1].height = 40
        ws.row_dimensions[r1 + 1].height = 25
        ws.row_dimensions[r1 + 2].height = 25

        current_row = r1 + 4  # 下一个部件从下一行开始（中间空两行）

    return current_row
def process_upper_sheet(wb, df_bridge, rule_file_path, safe_bridge_name):
    """
    处理上部结构 sheet，写入 wb。
    """
    ws_upper = wb.create_sheet("上部结构")

    df_upper = df_bridge[df_bridge["部位类型"] == "上部结构"]
    print(f"  [调试] df_upper 行数: {len(df_upper)}")

    if df_upper.empty:
        ws_upper.cell(row=1, column=1, value="无上部结构数据")
        print(f"  ⚠️ [{safe_bridge_name}] 无上部结构数据。")
        return

    try:
        # 1. 层级分组取最大扣分
        df_upper_result = traverse_bridge_data(df_upper)
        print(f"  [调试] traverse 后行数: {len(df_upper_result)}")

        if df_upper_result.empty:
            ws_upper.cell(row=1, column=1, value="上部结构无有效记录")
            print(f"  ⚠️ [{safe_bridge_name}] 上部结构无有效记录。")
            return

        # 2. 解析模板规则
        df_upper_rule_map = parse_template_rules(rule_file_path, "上部结构评分")

        # 3. 反推程度（使用 UPPER_COMPONENT_ALIAS）
        df_upper_result["程度"] = df_upper_result.apply(
            lambda r: score_to_degree(
                r["扣分分数"],
                df_upper_rule_map.get(
                    (UPPER_COMPONENT_ALIAS.get(r["部件类型"], r["部件类型"]), r["损坏类型"]),
                    {}
                )
            ),
            axis=1,
        )

        # 4. 补全缺失损坏类型（使用 UPPER_PART_DEFECT_TYPES 和 UPPER_COMPONENT_ALIAS）
        df_upper_result = fill_missing_defects(
            df_upper_result, UPPER_PART_DEFECT_TYPES, UPPER_COMPONENT_ALIAS
        )

        # 5. 汇总：按桥 + 部件 + 损坏类型取最大扣分那行
        df_upper_summary = (
            df_upper_result
            .sort_values("扣分分数", ascending=False)
            .groupby(["桥梁名称", "部件类型", "损坏类型"], as_index=False)
            .first()
        )

        # 6. 画表
        draw_upper_structure(ws_upper, df_upper_summary)
        print(f"  ✅ 上部结构处理成功。")
    except Exception as e:
        print(f"  ❌ 上部结构处理失败: {e}")
        ws_upper.cell(row=1, column=1, value=f"上部结构处理失败: {e}")

# if __name__ == "__main__":
#     # 1. 核心路径配置
#     raw_defect_excel_path = r"E:\Python\BridgeDefectAnalyzer\input\病害列表导出2026-10-09_modify1024.xlsx"
#     rule_file_path = r"E:\Python\BridgeDefectAnalyzer\input\城市桥梁评分计算.xlsx"
#     output_dir = r"E:\Python\BridgeDefectAnalyzer\output"

#     # 确保输出目录存在
#     os.makedirs(output_dir, exist_ok=True)

#     print(f"1. 正在读取全量病害总表: {raw_defect_excel_path}")
#     df_all = read_excel(raw_defect_excel_path, skip_hidden_rows=True, skip_hidden_cols=True)
#     # 提取上部结构数据
#     df_upper = df_all[df_all["部位类型"] == "上部结构"].copy()
#     # print(f"上部结构记录数: {len(df_upper)}")
#     # print(df_upper)
#     # # ===== 每个部件类型下的 损坏类型 → 病害类型 =====
#     # df_upper_valid = df_upper[
#     #     df_upper["病害类型"].notna()
#     #     & (df_upper["病害类型"].astype(str).str.strip() != "无明显病害")
#     # ].copy()

#     # print("\n" + "=" * 60)
#     # print("每个部件类型下的 损坏类型 → 病害类型")
#     # print("=" * 60)

#     # for comp_type, grp in df_upper_valid.groupby("部件类型"):
#     #     print(f"\n部件类型: {comp_type}")
#     #     for damage_type, sub in grp.groupby("损坏类型"):
#     #         defect_types = (
#     #             sub["病害类型"].dropna().astype(str).str.strip()
#     #             .replace("", pd.NA).dropna().unique().tolist()
#     #         )
#     #         print(f"  损坏类型: {damage_type}")
#     #         print(f"    病害类型（{len(defect_types)}种）: {defect_types}")
#     df_result = traverse_bridge_data(df_all)
#     defect_rule_map =parse_template_rules(rule_file_path, "上部结构评分")
#     df_result["程度"] = df_result.apply(
#     lambda r: score_to_degree(
#         r["扣分分数"],
#         defect_rule_map.get(
#             (UPPER_COMPONENT_ALIAS.get(r["部件类型"], r["部件类型"]), r["损坏类型"]),
#             {}
#         )
#     ),
#     axis=1,
#     )
#     pd.set_option('display.max_columns', None)  # 显示所有列
#     pd.set_option('display.width', 1000)        # 设置显示宽度，防止折行
#     pd.set_option('display.max_colwidth', 30)   # 列内容最大宽度

#     print("\n打印全部列的前 30 条数据：")
#     print(df_result.head(30).to_string(index=False))
#     df_result = fill_missing_defects(df_result, UPPER_COMPONENT_DEFECT_RULES, UPPER_COMPONENT_ALIAS)
#     # 1. 汇总时加上孔跨（防止不同孔跨的数据被合并）
#     df_upper_summary = (
#         df_result
#         .sort_values("扣分分数", ascending=False)
#         .groupby(["桥梁名称", "孔跨", "部件类型", "损坏类型"], as_index=False)
#         .first()
#     )

#     print("汇总后的数据：")
#     print(df_upper_summary.to_string(index=False))

#     # 按桥生成文件
#     for bridge, df_bridge in df_upper_summary.groupby("桥梁名称"):
#         wb = Workbook()
#         ws_upper = wb.active
#         ws_upper.title = "上部结构"
        
#         # 传当前桥汇总后的数据
#         draw_upper_structure(ws_upper, df_bridge)  
        
#         # 保存
#         safe_bridge = str(bridge).replace("/", "_").replace("\\", "_")
#         wb.save(os.path.join(output_dir, f"{safe_bridge}_上部结构评分表.xlsx"))
#         print(f"✅ 已生成: {safe_bridge}_上部结构评分表.xlsx")
