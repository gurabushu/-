from pytrends.request import TrendReq
import pandas as pd
import os
import subprocess
from datetime import datetime, timedelta
from datetime import datetime, timedelta

# 今日の日付を取得
end_date = datetime.today().strftime("%Y-%m-%d")

# 3年前の日付を取得
start_date = (datetime.today() - timedelta(days=3*365)).strftime("%Y-%m-%d")

# Googleトレンドの期間指定（3年間）
timeframe = f"{start_date} {end_date}"




#グーグルトレンドAPIに接続
pytrends = TrendReq(hl="ja-JP",tz=540) #日本語、日本時間

keywords = ["美容","ヘア","美容院","スタイリング"]

pytrends.build_payload(keywords, cat=0, timeframe=timeframe, geo="JP", gprop="")
trend_data = pytrends.interest_over_time()

if trend_data.empty:
    print(f"⚠️ {keywords} のデータが見つかりませんでした。")
    exit()

trend_data = trend_data.drop(columns=["isPartial"])

#日付を明示
trend_data.reset_index(inplace=True)
trend_data["date"] = pd.to_datetime(trend_data["date"])


print(trend_data.head())



df = pd.DataFrame(trend_data)

excel_path = "試用スクレイピング.xlsx"

with pd.ExcelWriter(excel_path,mode="a",engine="openpyxl") as writer:
    df.to_excel(writer,sheet_name="新しいシート", index=False)

print(f"データを{excel_path}の「新しいシート」に保存しました。")

subprocess.Popen(["start",excel_path], shell=True) 


