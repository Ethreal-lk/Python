import os 
import pandas as pd


def read_excel(file_path, sheet_name = 0):
    """读取指定 Sheet"""
    if not os.path.exists(file_path):
        raise FileNotFoundError("Excel 文件不存在: {}".format(file_path))
    
    return pd.read_excel(
        file_path,
        sheet_name=sheet_name
    )

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