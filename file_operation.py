import pandas as pd
from pptx import Presentation
import os
from datetime import datetime
import glob

def get_excel_file_paths(folder_with_date):
    excel_files = glob.glob(os.path.join(folder_with_date, '*.xlsx')) + glob.glob(os.path.join(folder_with_date, '*.xls'))
    return excel_files

def get_all_sheets(excel_file_path):
    all_sheets = {}
    xlsx = pd.ExcelFile(excel_file_path)
    sheet_names = xlsx.sheet_names
    for index, shhet_name in enumerate(sheet_names):
        all_sheets["sheet" + str(index + 1)] = pd.read_excel(xlsx, shhet_name)
    return all_sheets

def read_template_presentation():
    #ファイルパスを記述する
    return Presentation("")

def save_excel_file(prs, folder_with_date, monthly, agent_code, agent_name):
    extension = ".pptx"
    ppt_file_path = os.path.join(folder_with_date, monthly + "_AG定例会議_" + agent_name + "（" + agent_code + "）" + extension)
    prs.save(ppt_file_path)

if __name__ == "__main__":
    get_excel_file_paths()
    get_all_sheets()
    read_template_presentation()
    save_excel_file()
