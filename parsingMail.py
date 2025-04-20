import os
import re
import pandas as pd
from bs4 import BeautifulSoup

mail_dir = "./mails"

def parse_html(filepath, date):
    result = []

    new_headers = ["成交時間", "委託單號", "股號", "股票名稱", "類別", "股數", "單價", "價金"]
    old_headrs = new_headers + ["來源別"]

    with open(filepath, "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f, "html.parser")
        tables = soup.find_all('table')

    target_table = None
    for idx, table in enumerate(tables):
        print(f"parsing the {idx}th table of {filepath}")
        thead = table.find("thead", recursive=False)
        if not thead:
            print("thead not found")
            continue

        tr = thead.find("tr", recursive=False)
        if not tr:
            print("tr not found")
            continue

        th_texts = [td.get_text(strip=True) for td in tr.find_all("td", recursive=False)]
        
        # find the target table
        if th_texts == new_headers or th_texts == old_headrs:
            target_table = table
            break
    
    if target_table:
        for tr in target_table.find('tbody').find_all('tr'):
            cols = [td.get_text(strip=True) for td in tr.find_all('td')]
            if len(cols) >= 9:
                # 只取前9欄，避免超出 DataFrame 定義
                row = [date] + cols[:8]  # 前面加上成交日期
                result.append(row)
            else:
                print(f"⚠️ 資料欄位不足，跳過：{cols}")
    else:
            print("找不到成交資訊表格。")
    
    return result

def read_emails():
    result = []
    if not os.path.exists(mail_dir):
        print(f"❌ 資料夾不存在：{mail_dir}")
    for filename in os.listdir(mail_dir):
        if filename.endswith(".html"):
            path = os.path.join(mail_dir, filename)
            date_match = re.search(r"(\d{4})(\d{2})(\d{2})", filename)
            if date_match:
                y, m, d = date_match.groups()
                date_str = f"{y}/{m}/{d}"
                result.extend(parse_html(path, date_str))
    return result

def export_to_excel(data):
    export_columns = ["成交日期", "成交時間", "委託單號", "股號", "股票名稱", "類別", "股數", "單價", "價金"]
    df = pd.DataFrame(data, columns=export_columns)
    if not df.empty:
        df.to_excel("交易明細總表.xlsx", index=False)

    summarized_result = []
    transaction_type = {"買": 1, "賣": -1}
    for type_keyword, type_value in transaction_type.items():
        group_df = df[df["類別"].str.contains(type_keyword)]
        group_df["股數"] = pd.to_numeric(group_df["股數"].str.replace(',', ''), errors='coerce').fillna(0).astype(int)
        group_df["價金"] = pd.to_numeric(group_df["價金"].str.replace(',', ''), errors='coerce').fillna(0).astype(float)

        grouped = group_df.groupby(["股號", "股票名稱"]).agg({
            "股數": "sum",
            "價金": "sum"
        }).reset_index()

        grouped["平均成本"] = grouped["價金"] / grouped["股數"]
        grouped["總成本"] = grouped["價金"]
        grouped["方向"] = type_value
        grouped["類別"] = "現" + type_keyword
        summarized_result.append(grouped)

    total_stats = pd.concat(summarized_result, ignore_index=True)
    total_stats["方向加權股數"] = total_stats["股數"] * total_stats["方向"]
    net_shares_table = total_stats.groupby(["股號", "股票名稱"], as_index=False)["方向加權股數"].sum()
    net_shares_table.rename(columns={"方向加權股數": "淨股數"}, inplace=True)

    total_stats = total_stats.drop(columns=["方向", "方向加權股數"])
    total_stats = pd.merge(total_stats, net_shares_table, on=["股號", "股票名稱"], how="left")

    total_stats.to_excel("股票買賣成本統計.xlsx", index=False)

    print("✅ 成交資料已擷取並產出 Excel！")


def main():
    #scan all mails in the directories
    transactions = read_emails()
    export_to_excel(transactions)
    

main()