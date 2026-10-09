
from core.excel_handler import read_excel
from core.calculator import calculate_deck_max_dp
import sys

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

virus_file_path = r"G:\Python_lk\Python\BridgeDefectAnalyzer\input\virus.xlsx"
rule_file_path = r"G:\Python_lk\Python\BridgeDefectAnalyzer\input\城市桥梁评分计算.xlsx"


def main():
    # 读取 Excel 表格
    df = read_excel(virus_file_path, skip_hidden_rows=True)

    print("Excel 数据读取成功")
    print("数据行数：", len(df))

    # 计算桥面系各个损坏类型的最大扣分
    result = calculate_deck_max_dp(df)
    print(result)
    print(type(result))

    # print("\n桥面系各损坏类型最大扣分结果：")
    # print(result.to_string(index=False))


if __name__ == "__main__":
    main()