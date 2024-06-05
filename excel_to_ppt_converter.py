import number_transfer_to_textbox, number_transfer_to_table, text_inserter, file_operation, eigenvalues_from_excel, text_creater
import os
from datetime import datetime

def main():
    today_date = datetime.now().strftime("%Y_%m")
    #ファイルパスを記述する
    folder_path = ""
    folder_with_date = os.path.join(folder_path, today_date)
    excel_file_paths = file_operation.get_excel_file_paths(folder_with_date)

    for excel_file_path in excel_file_paths:
        all_sheets = file_operation.get_all_sheets(excel_file_path)
        monthly, agent_code, agent_name, period = eigenvalues_from_excel.get_eigenvalues(all_sheets["sheet1"])
        period_and_monthly = text_creater.integrate_period_and_monthly(period, monthly)
        period_and_cumulative = text_creater.integrate_period_and_cumulative(period)
        ppt_titles = text_creater.get_ppt_titles(monthly, agent_name)

        # テンプレプレゼンファイルの読み込み
        prs = file_operation.read_template_presentation()

        text_inserter.insert_ppt_titles(prs, 0, ppt_titles)

        #スライドタイトルの取得
        slide_titles = [
            text_creater.get_performance_by_period_and_monthly(period_and_monthly),
            text_creater.get_performance_by_period_and_monthly(period_and_cumulative),
            text_creater.get_customer_number_and_unitprice_by_monthly(period_and_monthly, "クラブ"),
            text_creater.get_customer_number_and_unitprice_by_cumulative(period_and_cumulative, "クラブ"),
            text_creater.get_customer_number_and_unitprice_by_categoly(monthly, "クラブ"),
            "実績（累計）カテゴリー別客数/単価（クラブ）",
            "実績 新規エージェント経由開拓（クラブ）",
            text_creater.get_new_and_existing_performance(monthly, "クラブ"),
            "実績（累計）新規/既存（クラブ）",
            text_creater.get_customer_number_and_unitprice_by_monthly(period_and_monthly, "SOL"),
            text_creater.get_customer_number_and_unitprice_by_cumulative(period_and_cumulative, "SOL"),
            text_creater.get_customer_number_and_unitprice_by_categoly(monthly, "SOL"),
            "実績（累計）カテゴリー別客数/単価（SOL）",
            "実績 新規開拓（SOL）",
            text_creater.get_new_and_existing_performance(monthly, "SOL"),
            "実績（累計）新規/既存（SOL）"
        ]

        for index, title in enumerate(slide_titles):
            text_inserter.insert_slide_titie(prs, index + 1, title)

        #スライドサブタイトルの挿入
        text_inserter.insert_sub_title(prs, 1, text_creater.get_performance_by_underbar(monthly), 1.55, 0.92)
        text_inserter.insert_sub_title(prs, 3, text_creater.get_performance(monthly), 2.24, 0.76)
        text_inserter.insert_sub_title(prs, 5, text_creater.get_performance(monthly), 1.70, 0.92)
        text_inserter.insert_sub_title(prs, 7, text_creater.get_performance(monthly), 2.25, 1.25)
        text_inserter.insert_sub_title(prs, 7, text_creater.get_performance(period_and_cumulative), 4.89, 1.25)
        text_inserter.insert_sub_title(prs, 8, text_creater.get_performance(monthly), 1.77, 0.92)
        text_inserter.insert_sub_title(prs, 10, text_creater.get_performance(monthly), 1.85, 0.96)
        text_inserter.insert_sub_title(prs, 12, text_creater.get_performance(monthly), 1.70, 0.92)
        text_inserter.insert_sub_title(prs, 14, text_creater.get_performance(monthly), 2.25, 1.25)
        text_inserter.insert_sub_title(prs, 14, text_creater.get_performance(period_and_cumulative), 4.89, 1.25)
        text_inserter.insert_sub_title(prs, 15, text_creater.get_performance(monthly), 1.86, 0.92)

        # #x月度実績の転記
        number_transfer_to_textbox.get_month_performances(prs, 1, all_sheets["sheet1"])
        number_transfer_to_textbox.get_month_performances(prs, 2, all_sheets["sheet2"])

        # #前期比の転記
        number_transfer_to_textbox.get_round_rates(prs, 1, all_sheets["sheet1"])
        number_transfer_to_textbox.get_round_rates(prs, 2, all_sheets["sheet2"])

        # #客数の転記
        number_transfer_to_textbox.get_club_customer_number_rate(prs, 3, all_sheets["sheet1"], 14, 16)
        number_transfer_to_textbox.get_club_customer_number_rate(prs, 10, all_sheets["sheet1"], 41, 43)
        number_transfer_to_textbox.get_club_customer_number_rate(prs, 4, all_sheets["sheet2"], 14, 16)
        number_transfer_to_textbox.get_club_customer_number_rate(prs, 11, all_sheets["sheet2"], 41, 43)

        # #新規開拓の転記
        number_transfer_to_textbox.get_new_development(prs, 7, all_sheets["sheet1"], 67, 69, 2, 0)
        number_transfer_to_textbox.get_new_development(prs, 7, all_sheets["sheet1"], 71, 73, 2, 1)
        number_transfer_to_textbox.get_new_development(prs, 14, all_sheets["sheet1"], 67, 69, 7, 0)
        number_transfer_to_textbox.get_new_development(prs, 14, all_sheets["sheet1"], 71, 73, 7, 1)

        # #新規開拓前期比の転記
        number_transfer_to_textbox.get_new_development_early_period_rate(prs, 7, all_sheets["sheet1"], 67, 69, 4, 0)
        number_transfer_to_textbox.get_new_development_early_period_rate(prs, 7, all_sheets["sheet1"], 71, 73, 4, 1)
        number_transfer_to_textbox.get_new_development_early_period_rate(prs, 14, all_sheets["sheet1"], 67, 69, 9, 0)
        number_transfer_to_textbox.get_new_development_early_period_rate(prs, 14, all_sheets["sheet1"], 71, 73, 9, 1)

        # #新規/既存の転記
        number_transfer_to_table.get_new_and_existing_month_club_performance(prs, 8, all_sheets["sheet1"], 32)
        number_transfer_to_table.get_new_and_existing_month_club_performance(prs, 9, all_sheets["sheet2"], 32)
        number_transfer_to_table.get_new_and_existing_month_club_performance(prs, 15, all_sheets["sheet1"], 59)
        number_transfer_to_table.get_new_and_existing_month_club_performance(prs, 16, all_sheets["sheet2"], 59)

        # #カテゴリー別客数/単価の転記
        number_transfer_to_table.get_customer_number_and_unitprice_by_category(prs, 5, all_sheets["sheet1"], 21)
        number_transfer_to_table.get_customer_number_and_unitprice_by_category(prs, 6, all_sheets["sheet2"], 21)
        number_transfer_to_table.get_customer_number_and_unitprice_by_category(prs, 12, all_sheets["sheet1"], 48)
        number_transfer_to_table.get_customer_number_and_unitprice_by_category(prs, 13, all_sheets["sheet2"], 48)

        # #地域別実績
        number_transfer_to_table.get_monthly_regional_performance(prs, 1, all_sheets["sheet3"])
        number_transfer_to_table.get_monthly_regional_performance(prs, 2, all_sheets["sheet4"])

        # #ファイル名を作成して保存する
        file_operation.save_excel_file(prs, folder_with_date, monthly, agent_code, agent_name)

if __name__ == "__main__":
    main()
