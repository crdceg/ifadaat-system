import pandas as pd
from datetime import datetime

# ====== اللغة ======
lang_path = "لغة.xlsx"
lang_new = "databases/language.xlsx"

df_lang = pd.read_excel(lang_path)

df_lang.rename(columns={
    "CODE_L": "CODE"
}, inplace=True)

df_lang["DATE"] = datetime.now().strftime("%d/%m/%Y")

df_lang = df_lang[["CODE","RESEARCHER","DEGREE","DEPARTMENT","TITLE","DATE"]]

df_lang.to_excel(lang_new, index=False)


# ====== الإحصاء ======
stat_path = "إحصاء.xlsx"
stat_new = "databases/statistics.xlsx"

df_stat = pd.read_excel(stat_path)

df_stat.rename(columns={
    "CODE_S": "CODE"
}, inplace=True)

df_stat["DATE"] = datetime.now().strftime("%d/%m/%Y")

df_stat = df_stat[["CODE","RESEARCHER","DEGREE","DEPARTMENT","TITLE","DATE"]]

df_stat.to_excel(stat_new, index=False)

print("✅ تم تجهيز الشيتات 100%")