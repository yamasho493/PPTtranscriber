def get_eigenvalues(sheet):
    sheet1_column_name_list = sheet.columns.tolist()
    monthly = sheet1_column_name_list[1][5:]
    agent_code = sheet.iloc[1, 1]
    agent_name = sheet.iloc[1, 3]
    period = sheet.iloc[19, 2][:3]
    return monthly, agent_code, agent_name, period

if __name__ == "__main__":
    get_eigenvalues()
