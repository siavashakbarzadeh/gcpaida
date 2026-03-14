import os
import pandas as pd
import glob

# مسیر پوشه حاوی فایل‌های csv سال 2022
folder_path = r"c:\Users\Utente\Documents\GitHub\gcpaida\2022_city_csv"

# دریافت لیست تمام فایل‌های csv داخل پوشه
csv_files = glob.glob(os.path.join(folder_path, "*.csv"))

# ایجاد یک لیست خالی برای ذخیره دیتافریم‌های هر شهر
dfs = []

for file in csv_files:
    # خواندن فایل
    df = pd.read_csv(file)
    
    # دریافت نام شهر از ستون 'denominazione_provincia' در ردیف اول
    if not df.empty and 'denominazione_provincia' in df.columns:
        city_name = df['denominazione_provincia'].iloc[0].strip().replace(" ", "_")
        
        # ساخت نام ستون جدید مثلا roma_infected
        col_name = f"{city_name.lower()}_infected"
        
        # انتخاب ستون‌های زمان و تعداد موارد
        df_subset = df[['data', 'totale_casi']].copy()
        
        # تغییر نام ستون 'totale_casi' به نام مدنظر
        df_subset.rename(columns={'totale_casi': col_name}, inplace=True)
        
        # تبدیل ستون data به فرمت datetime
        df_subset['data'] = pd.to_datetime(df_subset['data'])
        
        # ایندکس کردن بر اساس data
        df_subset.set_index('data', inplace=True)
        
        dfs.append(df_subset)

# ادغام همه دیتافریم‌ها بر اساس ایندکس زمان (data)
merged_df = pd.concat(dfs, axis=1, join='outer')

# برگرداندن data از حالت ایندکس به یک ستون عادی
merged_df.reset_index(inplace=True)

# مرتب‌سازی خطوط بر اساس زمان
merged_df.sort_values(by='data', inplace=True)

# مسیر فایل خروجی نهایی
output_path = os.path.join(folder_path, "merged_cities_2022.csv")

# ذخیره خروجی
merged_df.to_csv(output_path, index=False)

print(f"Merged successfully! Output saved to:\n{output_path}")
print(f"Number of rows: {len(merged_df)}")
print(f"Number of columns: {len(merged_df.columns)}")
print(f"Column names: {list(merged_df.columns)}")
