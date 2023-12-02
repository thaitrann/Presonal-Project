import os
import pandas as pd

def delete_files_in_directory(df):

    list_df_mail_columns = df.columns.to_list()
    for index, row in df.iterrows():
        code = row[list_df_mail_columns[0]]  # row from Code Sales
        name = row[list_df_mail_columns[1]]  # row from Name Sales
        directory = r"C:\Users\thomas.thai\Downloads\automail\Data sending\Data gửi sale\{} - {}".format(name, code)

        # check directory exists
        if not os.path.isdir(directory):
            print(f"Folder '{directory}' not exists.")
            continue

        # get file name in directory
        for filename in os.listdir(directory):
            file_path = os.path.join(directory, filename)
            try:
                if os.path.isfile(file_path):
                    os.remove(file_path)
                    print(f"Deleted file: '{filename}'.")
            except Exception as e:
                print(f"ERROR deleted file!: '{filename}': {str(e)}")

#list mail test
# df_mail_rsd = pd.read_excel(r'C:\Users\thomas.thai\Downloads\automail\Danh sách gửi mail\ds_mail.xlsx', sheet_name = 'RSD')
# df_mail_asm = pd.read_excel(r'C:\Users\thomas.thai\Downloads\automail\Danh sách gửi mail\ds_mail.xlsx', sheet_name = 'ASM')
# df_mail_ss = pd.read_excel(r'C:\Users\thomas.thai\Downloads\automail\Danh sách gửi mail\ds_mail.xlsx', sheet_name = 'SS')
# df_mail_sr = pd.read_excel(r'C:\Users\thomas.thai\Downloads\automail\Danh sách gửi mail\ds_mail.xlsx', sheet_name = 'SR')

#list mail final
df_mail_rsd = pd.read_excel(r'C:\Users\thomas.thai\Downloads\automail\Danh sách gửi mail\ds_mail_final.xlsx', sheet_name = 'RSD')
df_mail_asm = pd.read_excel(r'C:\Users\thomas.thai\Downloads\automail\Danh sách gửi mail\ds_mail_final.xlsx', sheet_name = 'ASM')
df_mail_ss = pd.read_excel(r'C:\Users\thomas.thai\Downloads\automail\Danh sách gửi mail\ds_mail_final.xlsx', sheet_name = 'SS')
df_mail_sr = pd.read_excel(r'C:\Users\thomas.thai\Downloads\automail\Danh sách gửi mail\ds_mail_final.xlsx', sheet_name = 'SR')

df_mail_rsd = df_mail_rsd[['Mã RSD 2023', 'Sales Manager']]
df_mail_asm = df_mail_asm[['Mã ASM 2023', 'ASM']]
df_mail_ss = df_mail_ss[['Mã sales sup 2023', 'Sales Sup']]
df_mail_sr = df_mail_sr[['Mã Sales Rep 2023', 'Sales Rep']]

delete_files_in_directory(df_mail_rsd)
delete_files_in_directory(df_mail_asm)
delete_files_in_directory(df_mail_ss)
delete_files_in_directory(df_mail_sr)

