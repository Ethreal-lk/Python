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


def calculate_bcim(scores, standard_weights, actual_weights):
    """
    计算 BCIm（桥面综合评分）

    参数：
        scores:
            各桥面要素评分，例如：
            [46.70, 70.00, 85.00, 70.00, 85.00, 100.00]

        standard_weights:
            规范权重，例如：
            [0.30, 0.10, 0.25, 0.10, 0.15, 0.10]

        actual_weights:
            实际参与权重，例如：
            [0.30, 0.10, 0.25, 0.10, 0.00, 0.00]

    返回：
        BCIm
    """

    # 实际参与权重之和
    total_actual_weight = sum(actual_weights)

    if total_actual_weight <= 0:
        return 100.0

    # 计算 BCIm
    bcim = 0.0

    for score, standard_weight, actual_weight in zip(
        scores,
        standard_weights,
        actual_weights
    ):

        # 不参与的桥面要素不计算
        if actual_weight <= 0:
            continue

        # 重分配权重
        redistributed_weight = (
            standard_weight / total_actual_weight
        )

        bcim += score * redistributed_weight

    return bcim


def calculate_bsim(scores, actual_weights):
    """
    计算 BSIm

    实际参与的桥面要素取最低评分。
    未参与的桥面要素按 100 分处理。

    参数：
        scores:
            各桥面要素评分

        actual_weights:
            各桥面要素实际参与权重

    返回：
        BSIm
    """

    if not scores:
        return 100.0

    adjusted_scores = []

    for score, actual_weight in zip(
        scores,
        actual_weights
    ):

        if actual_weight > 0:
            adjusted_scores.append(score)
        else:
            adjusted_scores.append(100.0)

    return min(adjusted_scores)