
import pandas as pd
from bs4 import BeautifulSoup

# File paths
table_files = [
    r"C:\Business_Quant\Business Quant Dataset - Html Tables\table_9.html",
    r"C:\Business_Quant\Business Quant Dataset - Html Tables\table_12.html",
    r"C:\Business_Quant\Business Quant Dataset - Html Tables\table_62.html"
]

# Function to extract tables from HTML
def extract_tables(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        soup = BeautifulSoup(file, 'html.parser')
        tables = soup.find_all('table')
        df_list = []
        for table in tables:
            df = pd.read_html(str(table))[0]
            df_list.append(df)
        return df_list

# Extract tables from each file
all_tables = {}
for i, file in enumerate(table_files, 1):
    all_tables[f'table_{i}'] = extract_tables(file)

# Display the extracted tables
for table_name, tables in all_tables.items():
    print(f"\n{table_name}")
    for idx, df in enumerate(tables):
        print(f"Table {idx + 1}:")
        print(df.head())

# Save tables to an Excel file with proper formatting
with pd.ExcelWriter(r'C:\Business_Quant\formatted_output.xlsx', engine='xlsxwriter') as writer:
    for table_name, tables in all_tables.items():
        for idx, df in enumerate(tables):
            sheet_name = f"{table_name}_part_{idx + 1}"
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value)
            worksheet.autofilter(0, 0, len(df.index), len(df.columns) - 1)

print("Formatted Excel file saved as C:\Business_Quant\formatted_output.xlsx")