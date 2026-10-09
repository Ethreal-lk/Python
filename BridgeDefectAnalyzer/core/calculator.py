import pandas as pd
from  core.defect_rules import DECK_COMPONENT_DEFECT_RULES


def calculate_component_score(component_df, component_name):
    """
    计算桥面要素评分

    参数：
        component_df：
            当前桥面要素的数据，类型为 pandas.DataFrame

        component_name：
            桥面要素名称，例如：
            "桥面铺装"
            "伸缩装置"
            "栏杆、护栏"
            "排水系统"

    计算过程：
        每个“损坏类型”取其对应病害中的最大扣分分数。

        μij = DPij / ΣDPij
        ωij = 3μij³ - 5.5μij² + 3.5μij
        DPi×ωij = DPij × ωij

        最终评分 = 100 - Σ(DPi×ωij)
    """

    df = component_df.copy()

    # --------------------------------------------------------
    # 1. 检查部件规则是否存在
    # --------------------------------------------------------
    if component_name not in DECK_COMPONENT_DEFECT_RULES:
        raise ValueError(
            f"没有找到部件 [{component_name}] 的病害规则"
        )

    component_rules = DECK_COMPONENT_DEFECT_RULES[
        component_name
    ]

    # --------------------------------------------------------
    # 2. 扣分分数转数字
    # --------------------------------------------------------
    df["扣分分数"] = pd.to_numeric(
        df["扣分分数"],
        errors="coerce"
    ).fillna(0)

    # --------------------------------------------------------
    # 3. 病害类型 → 损坏类型
    # --------------------------------------------------------
    def get_damage_category(defect_type):

        for category, defect_types in component_rules.items():

            if defect_type in defect_types:
                return category

        return None

    df["损坏类型"] = df["病害类型"].apply(
        get_damage_category
    )

    # --------------------------------------------------------
    # 4. 只保留已经归类且有扣分的病害
    # --------------------------------------------------------
    df = df[
        df["损坏类型"].notna()
        & (df["扣分分数"] > 0)
    ].copy()

    # --------------------------------------------------------
    # 5. 建立完整的损坏类型
    # --------------------------------------------------------
    damage_types = list(
        component_rules.keys()
    )

    result = pd.DataFrame({
        "损坏类型": damage_types
    })

    # --------------------------------------------------------
    # 6. 每个损坏类型取最大的扣分分数
    # --------------------------------------------------------
    max_dp = (
        df.groupby("损坏类型")["扣分分数"]
        .max()
        .rename("单项扣分DPij")
    )

    result = result.merge(
        max_dp,
        on="损坏类型",
        how="left"
    )

    result["单项扣分DPij"] = (
        result["单项扣分DPij"]
        .fillna(0)
    )

    # --------------------------------------------------------
    # 7. 总 DP
    # --------------------------------------------------------
    total_dp = result["单项扣分DPij"].sum()

    # --------------------------------------------------------
    # 8. 计算比重和权重
    # --------------------------------------------------------
    if total_dp > 0:

        result["比重μij"] = (
            result["单项扣分DPij"]
            / total_dp
        )

        u = result["比重μij"]

        result["权重ωij"] = (
            3 * u ** 3
            - 5.5 * u ** 2
            + 3.5 * u
        )

        result["DPi×ωij"] = (
            result["单项扣分DPij"]
            * result["权重ωij"]
        )

    else:

        result["比重μij"] = 0.0
        result["权重ωij"] = 0.0
        result["DPi×ωij"] = 0.0

    # --------------------------------------------------------
    # 9. 总扣分
    # --------------------------------------------------------
    total_deduction = result["DPi×ωij"].sum()

    # --------------------------------------------------------
    # 10. 最终评分
    # --------------------------------------------------------
    score = 100 - total_deduction

    return {
        "component_name": component_name,
        "score": score,
        "total_deduction": total_deduction,
        "details": result
    }


def calculate_deck_max_dp(df):
    """
    提取桥面系各部件、各损坏类型的最大扣分。

    当前支持：
        桥面铺装

    后续增加其他部件规则后，无需修改计算逻辑。
    """

    df = df.copy()

    # 1. 筛选桥面系数据
    df = df[
        df["部位类型"] == "桥面系"
    ].copy()

    # 2. 扣分分数转换为数字
    df["扣分分数"] = pd.to_numeric(
        df["扣分分数"],
        errors="coerce"
    ).fillna(0)

    results = []

    # 3. 遍历桥面系下的每个部件
    for component_name, component_rules in (
        DECK_COMPONENT_DEFECT_RULES.items()
    ):

        # 筛选当前部件
        component_df = df[
            df["部件类型"] == component_name
        ].copy()

        # 4. 遍历当前部件的每种损坏类型
        for damage_category, defect_types in (
            component_rules.items()
        ):

            # 找出属于当前损坏类型的原始病害
            damage_df = component_df[
                component_df["病害类型"].isin(defect_types)
                & (component_df["扣分分数"] > 0)
            ]

            # 5. 提取最大扣分；没有记录时记为0
            max_dp = (
                damage_df["扣分分数"].max()
                if not damage_df.empty
                else 0
            )

            results.append({
                "部件类型": component_name,
                "损坏类型": damage_category,
                "单项扣分DPij": max_dp,
            })

    # 6. 转换为结果表
    result = pd.DataFrame(results)

    return result