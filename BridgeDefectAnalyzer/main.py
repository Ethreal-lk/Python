from core.excel_handler import read_excel
import sys
sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

df = read_excel(r"G:\Python_lk\Python\BridgeDefectAnalyzer\input\virus.xlsx")
print("读取成功")
print("列名：")
print(df.columns.tolist())

print("\n第一行数据：")
print(df.iloc[2])