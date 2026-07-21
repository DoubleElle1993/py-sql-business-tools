import pandas as pd
import os

excel_folder_path = os.path.join(os.getcwd(), 'Csv_SD')

df_list = []
for excel_fle in os.listdir(excel_folder_path):
    if excel_fle.endswith('.xlsx'):
        file_path = os.path.join(excel_folder_path, excel_fle)
        df = pd.read_excel(file_path)
        df_list.append(df)

final_df = pd.concat(df_list, ignore_index=True)
output_path = os.path.join(excel_folder_path, 'name.xlsx')
final_df.to_excel(output_path, index=False)