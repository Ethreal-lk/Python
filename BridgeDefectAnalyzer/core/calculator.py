import pandas as pd
from  core.defect_rules import DECK_COMPONENT_DEFECT_RULES


def calculate_pavement_score(pavement_df):
    """
    计算桥面铺装评分

    DPij：
        每个“损坏类型”取其对应病害中的最大扣分分数。

    μij = DPij / ΣDPij
    ωij = 3μij³ - 5.5μij² + 3.5μij
    DPi×ωij = DPij × ωij
    最终评分 = 100 - Σ(DPi×ωij)
    """

    df = pavement_df.copy()

    # --------------------------------------------------------
    # 1. 扣分分数转数字
    # --------------------------------------------------------
    df["扣分分数"] = pd.to_numeric(
        df["扣分分数"],
        errors="coerce"
    ).fillna(0)

    # --------------------------------------------------------
    # 2. 病害类型 → 损坏类型
    # --------------------------------------------------------
    def get_damage_category(defect_type):
        for category, defect_types in DECK_COMPONENT_DEFECT_RULES["桥面铺装"].items():
            if defect_type in defect_types:
                return category

        return None

    df["损坏类型"] = df["病害类型"].apply(
        get_damage_category
    )

    # --------------------------------------------------------
    # 3. 只保留已经归类且有扣分的病害
    # --------------------------------------------------------
    df = df[
        df["损坏类型"].notna()
        & (df["扣分分数"] > 0)
    ].copy()

    # --------------------------------------------------------
    # 4. 建立完整的损坏类型
    # --------------------------------------------------------
    damage_types = list(
        DECK_COMPONENT_DEFECT_RULES["桥面铺装"].keys()
    )

    result = pd.DataFrame({
        "损坏类型": damage_types
    })

    # --------------------------------------------------------
    # 5. 每个损坏类型取最大的扣分分数
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

    result["单项扣分DPij"] = result["单项扣分DPij"].fillna(0)

    # --------------------------------------------------------
    # 6. 总 DP
    # --------------------------------------------------------
    total_dp = result["单项扣分DPij"].sum()

    # --------------------------------------------------------
    # 7. 计算比重和权重
    # --------------------------------------------------------
    if total_dp > 0:

        result["比重μij"] = (
            result["单项扣分DPij"] / total_dp
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
    # 8. 总扣分
    # --------------------------------------------------------
    total_deduction = result["DPi×ωij"].sum()

    # --------------------------------------------------------
    # 9. 最终评分
    # --------------------------------------------------------
    score = 100 - total_deduction

    return {
        "score": score,
        "total_deduction": total_deduction,
        "details": result
    }