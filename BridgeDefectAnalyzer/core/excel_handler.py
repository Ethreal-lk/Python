import os 
import pandas as pd


import os
import pandas as pd
from openpyxl import load_workbook


def read_excel(file_path, sheet_name=0, skip_hidden_rows=False):
    """读取指定 Sheet"""

    if not os.path.exists(file_path):
        raise FileNotFoundError(
            "Excel 文件不存在: {}".format(file_path)
        )

    # 不需要跳过隐藏行，直接使用 pandas
    if not skip_hidden_rows:
        return pd.read_excel(
            file_path,
            sheet_name=sheet_name
        )

    # 需要跳过隐藏行
    wb = load_workbook(file_path, data_only=True)

    ws = (
        wb.worksheets[sheet_name]
        if isinstance(sheet_name, int)
        else wb[sheet_name]
    )

    # 第一行作为表头
    headers = [cell.value for cell in ws[1]]

    # 读取可见数据行
    data = [
        [cell.value for cell in row]
        for row in ws.iter_rows(min_row=2)
        if not ws.row_dimensions[row[0].row].hidden
    ]

    return pd.DataFrame(data, columns=headers)

def get_sheet_names(file_path):
    """获取 Sheet 名称"""
    if not os.path.exists(file_path):
        raise FileNotFoundError("Excel 文件不存在: {}".format(file_path))

    excel_file = pd.ExcelFile(file_path)
    return excel_file.sheet_names

def write_excel(data, file_path, sheet_name="Sheet1"):
    """写入单个 Sheet"""
    data.to_excel(
        file_path,
        sheet_name=sheet_name,
        index=False
    )

def write_sheets(data_dict, file_path):
    """写入多个 Sheet"""

    with pd.ExcelWriter(
        file_path,
        engine="openpyxl"
    ) as writer:

        for sheet_name, data in data_dict.items():
            data.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )

def append_sheet(data, file_path, sheet_name):
    """追加 Sheet"""

    with pd.ExcelWriter(
        file_path,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace"
    ) as writer:

        data.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False
        )


def check_excel(file_path):
    """检查 Excel 是否可以正常读取"""

    try:
        get_sheet_names(file_path)
        return True
    except Exception:
        return False