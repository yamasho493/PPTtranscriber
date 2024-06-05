def integrate_period_and_monthly(period, monthly):
    return period + monthly

def integrate_period_and_cumulative(period):
    return period + "累計"

def get_ppt_titles(monthly, agent_name):
    titles = []
    title_of_first_line = monthly + "　定例会資料"
    title_of_second_line = "株式会社" + agent_name + "御中"
    titles.extend([title_of_first_line, title_of_second_line])
    return titles

def get_customer_number_and_unitprice_by_monthly(period_and_monthly, media):
    if media == "クラブ":
        return "実績（" + period_and_monthly + "）客数/単価（クラブ）" 
    elif media == "SOL":
        return "実績（" + period_and_monthly + "）客数/単価（SOL）"
    else:
        return ""
    
def get_customer_number_and_unitprice_by_cumulative(period_and_cumulative, media):
    if media == "クラブ":
        return "実績（" + period_and_cumulative + "）客数/単価（クラブ）" 
    elif media == "SOL":
        return "実績（" + period_and_cumulative + "）客数/単価（SOL）"
    else:
        return ""

def get_customer_number_and_unitprice_by_categoly(monthly, media):
    if media == "クラブ":
        return "実績（" + monthly + "）カテゴリー別客数/単価（クラブ）" 
    elif media == "SOL":
        return "実績（" + monthly + "）カテゴリー別客数/単価（SOL）"
    else:
        return ""
    
def get_new_and_existing_performance(monthly, media):
        if media == "クラブ":
            return "実績（" + monthly + "）新規/既存（クラブ）" 
        elif media == "SOL":
            return "実績（" + monthly + "）新規/既存（SOL）"
        else:
            return ""
        
def get_performance(text):
    return "■貴社" + text + "実績"

def get_performance_by_underbar(monthly):
    return "■貴社" + monthly + "＿実績"

def get_performance_by_period_and_monthly(text):
    return "実績報告（" + text + "）"

if __name__ == "__main__":
    integrate_period_and_monthly()
    integrate_period_and_cumulative()
    get_ppt_titles()
    get_customer_number_and_unitprice_by_monthly()
    get_customer_number_and_unitprice_by_cumulative()
    get_customer_number_and_unitprice_by_categoly()
    get_new_and_existing_performance()
    get_performance()
    get_performance_by_underbar()
    get_performance_by_period_and_monthly()
