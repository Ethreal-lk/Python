from core.excel_handler import read_excel
from core.calculator import calculate_component_score
from core.defect_rules import DECK_COMPONENT_TYPES
import sys

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")


file_path = r"G:\Python_lk\Python\BridgeDefectAnalyzer\input\virus.xlsx"


def analyze_deck(df, component_types):
    """
    分析所有桥梁的桥面系各部件评分

    分组层级：
        桥梁名称
        ↓
        部件类型
        ↓
        分幅
        ↓
        孔跨
    """

    for bridge_name, bridge_df in df.groupby("桥梁名称"):

        print("\n" + "=" * 60)
        print(f"桥梁：{bridge_name}")
        print("=" * 60)

        # ====================================================
        # 1. 筛选桥面系
        # ====================================================
        deck_df = bridge_df[
            bridge_df["部位类型"] == "桥面系"
        ]

        if deck_df.empty:
            print("没有桥面系数据，跳过")
            continue

        # ====================================================
        # 2. 遍历桥面系下面的每个部件
        # ====================================================
        for component_type in component_types:

            component_df = deck_df[
                deck_df["部件类型"] == component_type
            ]

            print("\n" + "-" * 50)
            print(f"部件：{component_type}")
            print(f"数据：{len(component_df)} 条")
            print("-" * 50)

            if component_df.empty:
                print("没有该部件数据，跳过")
                continue

            # =================================================
            # 3. 按 分幅 + 孔跨
            # =================================================
            for (fenfu, kongkua), span_df in component_df.groupby(
                ["分幅", "孔跨"]
            ):

                print(
                    f"\n分幅：{fenfu}    "
                    f"孔跨：{kongkua}"
                )

                print(
                    f"本孔跨数据：{len(span_df)} 条"
                )

                # =============================================
                # 4. 调用统一评分函数
                # =============================================
                result = calculate_component_score(
                    span_df,
                    component_type
                )

                # =============================================
                # 5. 输出结果
                # =============================================
                print(
                    f"总扣分："
                    f"{result['total_deduction']:.2f}"
                )

                print(
                    f"评分："
                    f"{result['score']:.2f}"
                )

                # =============================================
                # 6. 输出详细计算过程
                # =============================================
                if not result["details"].empty:

                    print("\n损坏类型计算明细：")

                    print(
                        result["details"][
                            [
                                "损坏类型",
                                "单项扣分DPij",
                                "比重μij",
                                "权重ωij",
                                "DPi×ωij"
                            ]
                        ].to_string(index=False)
                    )


if __name__ == "__main__":

    df = read_excel(
        file_path,
        skip_hidden_rows=True
    )

    print(
        f"Excel读取完成，共 {len(df)} 条数据"
    )

    analyze_deck(
        df,
        DECK_COMPONENT_TYPES
    )

if __name__ == "__main__":

    df = read_excel(
        file_path,
        skip_hidden_rows=True
    )

    print(f"Excel读取完成，共 {len(df)} 条数据")

    analyze_deck(
        df,
        DECK_COMPONENT_TYPES
    )