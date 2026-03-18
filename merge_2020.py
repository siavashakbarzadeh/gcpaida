import os
import pandas as pd
import glob

# مسیر پوشه حاوی فایل‌های csv سال 2020
folder_path = r"c:\Users\Utente\Documents\GitHub\gcpaida\2020\2020_city_csv"

# مسیر خروجی در پوشه 2020
output_dir = r"c:\Users\Utente\Documents\GitHub\gcpaida\2020"

# دریافت لیست تمام فایل‌های csv داخل پوشه
csv_files = glob.glob(os.path.join(folder_path, "*.csv"))

# ایجاد یک لیست خالی برای ذخیره دیتافریم‌های هر شهر
dfs = []

for file in csv_files:
    df = pd.read_csv(file)
    if not df.empty and 'denominazione_provincia' in df.columns:
        city_name = df['denominazione_provincia'].iloc[0].strip().replace(" ", "_")
        col_name = f"infected_{city_name}"
        df_subset = df[['data', 'totale_casi']].copy()
        df_subset.rename(columns={'totale_casi': col_name}, inplace=True)
        df_subset['data'] = pd.to_datetime(df_subset['data'])
        df_subset.set_index('data', inplace=True)
        dfs.append(df_subset)

merged_df = pd.concat(dfs, axis=1, join='outer')
merged_df.reset_index(inplace=True)
merged_df.sort_values(by='data', inplace=True)

output_path = os.path.join(output_dir, "merged_cities_2020.csv")
merged_df.to_csv(output_path, index=False)

print(f"Merged successfully! Output saved to:\n{output_path}")
print(f"Number of rows: {len(merged_df)}")
print(f"Number of columns: {len(merged_df.columns)}")
print(f"Column names: {list(merged_df.columns)}")
