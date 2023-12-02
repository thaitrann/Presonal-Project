import sys
import os
sys.path.insert(0, r'C:\Users\thomas.thai\Downloads\automail\automail_sale')
import pandas as pd
import time
from datetime import datetime
from mail.send_mail import send_mail
start_time = time.time()

#functions
# define rules
def check_condition(row):
    start_time = time.time()
    if row["Loài"] == "Aqua" and row["Thương hiệu"] == "Deheus" and row["A00_BusinessUnit"] == "FNA":
            return "Torin Trường tại FNA"
    
    elif pd.isna(row["Khu vực"]) and row["Loài"] != "Aqua" and row["Thương hiệu"] == "Deheus" \
            and row["A00_BusinessUnit"] == "PHP":
            return "Michel Minh tại PHP"
    
    elif pd.isna(row["Khu vực"]) and row["Loài"] != "Aqua" and row["Thương hiệu"] == "Deheus" \
            and row["A00_BusinessUnit"] == "FNA":
            return "Michel Minh tại FNA"
    
    elif row["A00_BusinessUnit"] == "PCCT":
            return "AQUA"
    
    elif row["Khu vực"] == "Farm NA":
            return "Farm NA"

def pivot_table_format(df):
    data_filter = df.rename(columns = {'Invoice date': 'Ngày lấy hàng', 'Target Customer': 'Tên khách hàng', 'Warehouse name': 'Tên kho',
                'Nhà máy': 'Công ty', 'Description': 'Tên sản phẩm', 'Quantity (kg)': 'Sản lượng (kg)', 'Loài': 'Nhóm', 'Loại con': 'Loài'})

    pivot_data_filter = data_filter.pivot_table(index = ['Mã khách hàng', 'Tên khách hàng', 
                                                'Tên kho', 'Công ty', 'Ngày lấy hàng', 'Mã sản phẩm', 'Tên sản phẩm', 'Nhóm',
                                                'Loài'], 
                                                values = 'Sản lượng (kg)', aggfunc = 'sum')
    df_pivot_sales = data_filter[['Mã khách hàng', 'Mã sản phẩm', 'Thương hiệu', 'Sales Rep', 'Sales Sup', 'ASM', 'Sales Manager']].drop_duplicates()
    df = pd.DataFrame(pivot_data_filter.to_records())
    df_pivot_merged = df.merge(df_pivot_sales, on = ['Mã khách hàng', 'Mã sản phẩm'], how = 'left')

    return df_pivot_merged

def send_data_sales_by_number(list_of_columns, list_of_numbers, final_df, max_date, designation):
    empty_data_filter_list = []

    for i in list_of_numbers:
        data_filter = final_df[(final_df[list_of_columns[1]] == i[0]) & (final_df[list_of_columns[2]] == i[1])]
        
        if data_filter.empty:
            empty_data_filter_list.append(i)
            continue

        data_filter = data_filter.rename(columns = {'Invoice date': 'Ngày lấy hàng', 'Target Customer': 'Tên khách hàng', 'Warehouse name': 'Tên kho',
                        'Nhà máy': 'Công ty', 'Description': 'Tên sản phẩm', 'Quantity (kg)': 'Sản lượng (kg)', 'Loài': 'Nhóm', 'Loại con': 'Loài'})
        
        pivot_data_filter = data_filter.pivot_table(index = ['Mã khách hàng', 'Tên khách hàng', 
                                                        'Tên kho', 'Công ty', 'Ngày lấy hàng', 'Mã sản phẩm', 'Tên sản phẩm', 'Nhóm',
                                                        'Loài'], 
                                                        values = 'Sản lượng (kg)', aggfunc = 'sum')
        df_pivot_sales = data_filter[['Mã khách hàng', 'Mã sản phẩm', 'Thương hiệu', 'Sales Rep', 'Sales Sup', 'ASM', 'Sales Manager']].drop_duplicates()

        df = pd.DataFrame(pivot_data_filter.to_records())
        df_pivot_merged = df.merge(df_pivot_sales, on = ['Mã khách hàng', 'Mã sản phẩm'], how = 'left')

        # #elements use for sending mail
        sum_quantity = df_pivot_merged['Sản lượng (kg)'].sum()
        number_ = data_filter[list_of_columns[1]].unique()
        name_ = data_filter[list_of_columns[2]].unique()
        mail_ = data_filter[list_of_columns[3]].unique()
        mail_cc = data_filter[list_of_columns[4]].unique()
        string_date = max_date.replace("-", "")

        end_time = time.time()
        print("---------- DONE FILTER DATA BY NUMBER! ----------")
        execution_time = end_time - start_time
        print("Execution time: ", execution_time)
        print("-------------------------------------")
        send_mail(number_, name_, mail_, mail_cc, string_date, df_pivot_merged, max_date, sum_quantity)

    if len(empty_data_filter_list) > 0:
        empty_data_filter_df = pd.DataFrame(empty_data_filter_list, columns=['number', 'Name', 'mail', 'mail_cc'])
        empty_data_filter_df.to_excel(r'C:\Users\thomas.thai\Downloads\automail\Data sending\danh sách không có sản lượng\empty_data_filters_{}.xlsx'\
                                      .format(designation), index=False)
    else:
        print('---------- NO EMPTY! ----------')
#sale data
php = pd.read_excel(r'C:\Users\thomas.thai\Downloads\automail\sales_data\Data\PHP.xlsx', header = 0)
mns = pd.read_excel(r'C:\Users\thomas.thai\Downloads\automail\sales_data\Data\MNS.xlsx', header = 0)
pbh = pd.read_excel(r'C:\Users\thomas.thai\Downloads\automail\sales_data\Data\PBH.xlsx', header = 0)
dhv = pd.read_excel(r'C:\Users\thomas.thai\Downloads\automail\sales_data\Data\DHV.xlsx', header = 0)
concat_df = pd.concat([php, mns, pbh, dhv], ignore_index = True)

#sales mapping info
mapping_all = pd.read_excel(r'C:\Users\thomas.thai\Downloads\Mapping\Mapping_automail.xlsx', \
                            header = 0, sheet_name = 'Mapping')
mapping_all_info_customer = mapping_all[['Khu vực', 'Mã khách hàng', 'Mã Sales Rep 2023']].drop_duplicates()

mapping_all_sale_info = mapping_all[['Mã Sales Rep 2023','Tên Sales Rep','Mã sales sup 2023', 'Tên sales sup', 'Mã ASM 2023','Tên ASM', 'Mã RSD 2023', 'Tên RSD']].drop_duplicates().dropna()

mapping_2_sales = pd.read_excel(r'C:\Users\thomas.thai\Downloads\Mapping\Mapping_automail.xlsx', \
                                header = 0, sheet_name = 'KH mapping 2 sales')
#FG
finished_goods = pd.read_excel(r'C:\Users\thomas.thai\Downloads\FG\Finished goods.xlsx', \
                               header = 0, sheet_name = 'FG')

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

#transforms
columns = ['Sales Order Number', 'VAT serial number',
    'VAT invoice number', 'SI No', 'Voucher no', 'Due Date',
    'Site', 'Warehouse', 'Location', 'Line number',
    'Search name', 'Item Group', 'CW quantity',
    'Unit', 'Sales price', 'Price include sales tax',
    'Currency', 'Discount Amount', 'Discount percent', 'MULTILINE DISCOUNT',
    'MULTINLINE DISCOUNT PERCENTAGE', 'Gross amount',
    'Total discount amount', 'Amount', 'Amount before tax', 'Tax amount',
    'Amount included sales tax', 'Exchange rate', 'Sales tax group',
    'Item sales tax group',
    'Salesman Code', 'Salesman Name',
    'A04_Location', 'A09_Species',
    'Sales Pool']

# #filter row and column not use
concat_df.drop(columns, inplace=True, axis=1)

filtered_df = concat_df.loc[(concat_df['Invoice type'] != 'Free Text Invoice')\
                     & (concat_df['A08_Division'] != 'PREMIX')\
                     & (concat_df['A00_BusinessUnit'] != 'FDN2')\
                     & (concat_df['Customer group'].isin(['LOC_EXT', 'FOR_EXT']))]

#add info Item base Item code
final_df = finished_goods[['Mã sản phẩm', 'Thương hiệu', 'Loài', 'Loại con']]\
    .merge(filtered_df, left_on = 'Mã sản phẩm', right_on = 'Item Code', how = 'right')\
    .merge(mapping_all_info_customer[['Khu vực', 'Mã khách hàng']], left_on='Target Customer Code', right_on='Mã khách hàng', how='left')


# #add info customer has 2 sales rep
final_df = final_df.merge(mapping_2_sales[['Mã khách hàng', 'Ghi chú', 'Mã Sales Rep 2023']], \
                    left_on = ['Target Customer Code', 'Thương hiệu'], \
                    right_on = ['Mã khách hàng', 'Ghi chú'], how = 'left')

# #fill blank column by mapping all sale rep
final_df = final_df.merge(mapping_all_info_customer[['Mã khách hàng', 'Mã Sales Rep 2023']], \
                    left_on = ['Target Customer Code'], 
                    right_on = ['Mã khách hàng'], how = 'left', suffixes=('', '_m_all'))
final_df['Mã Sales Rep 2023'] = final_df['Mã Sales Rep 2023'].fillna(final_df['Mã Sales Rep 2023_m_all'])
final_df = final_df.merge(mapping_all_sale_info, on = 'Mã Sales Rep 2023', how = 'left')

# #fill blank column by torin trường, michel minh,...
final_df["temp"] = final_df.apply(check_condition, axis=1)
final_df['Tên RSD'] = final_df['Tên RSD'].fillna(final_df['temp'])

final_df.loc[(final_df['Tên RSD'].notnull()) & (final_df['A00_BusinessUnit'] == 'PCCT'), 'Tên RSD'] = 'AQUA'
final_df.loc[(final_df['Tên ASM'].notnull()) & (final_df['A00_BusinessUnit'] == 'PCCT'), 'Tên ASM'] = 'AQUA'
final_df.loc[(final_df['Tên sales sup'].notnull()) & (final_df['A00_BusinessUnit'] == 'PCCT'), 'Tên sales sup'] = 'AQUA'
final_df.loc[(final_df['Tên Sales Rep'].notnull()) & (final_df['A00_BusinessUnit'] == 'PCCT'), 'Tên Sales Rep'] = 'AQUA'

# #rename column
final_df = final_df.rename(columns={"Tên RSD": "Sales Manager", 'A00_BusinessUnit': 'Nhà máy', 'Quantity': 'Quantity (kg)', 
                                    'Tên ASM': 'ASM', 'Tên Sales Rep': 'Sales Rep', 'Tên sales sup': 'Sales Sup'})

final_df['Mã Sales Rep 2023'] = final_df['Mã Sales Rep 2023'].fillna(0).astype(int)
final_df['Mã sales sup 2023'] = final_df['Mã sales sup 2023'].fillna(0).astype(int)
final_df['Mã ASM 2023'] = final_df['Mã ASM 2023'].fillna(0).astype(int)
final_df['Mã RSD 2023'] = final_df['Mã RSD 2023'].fillna(0).astype(int)

final_df['Invoice date'] = pd.to_datetime(final_df['Invoice date'])
max_date = final_df['Invoice date'].dt.date.max()
date_obj = datetime.strptime(str(max_date), '%Y-%m-%d')
max_date = date_obj.strftime('%d-%m-%Y')

# #merge by number rsd, asm, ss
final_df = final_df.merge(df_mail_rsd[['Mã RSD 2023', 'mail rsd', 'cc_rsd']], on = 'Mã RSD 2023', how = 'left')\
                .merge(df_mail_asm[['Mã ASM 2023', 'mail asm', 'cc_asm']], on = 'Mã ASM 2023', how = 'left')\
                .merge(df_mail_ss[['Mã sales sup 2023', 'mail ss', 'cc_ss']], on = 'Mã sales sup 2023', how = 'left')\
                .merge(df_mail_sr[['Mã Sales Rep 2023', 'mail sr', 'cc_sr']], on = 'Mã Sales Rep 2023', how = 'left')

list_of_rsd_numbers = df_mail_rsd[['Mã RSD 2023', 'Sales Manager', 'mail rsd', 'cc_rsd']].values.tolist()
list_of_asm_numbers = df_mail_asm[['Mã ASM 2023', 'ASM', 'mail asm', 'cc_asm']].values.tolist()
list_of_ss_numbers = df_mail_ss[['Mã sales sup 2023', 'Sales Sup', 'mail ss', 'cc_ss']].values.tolist()
list_of_sr_numbers = df_mail_sr[['Mã Sales Rep 2023', 'Sales Rep', 'mail sr', 'cc_sr']].values.tolist()

# #column use of sending mail
list_df_mail_rsd_columns = df_mail_rsd.columns.to_list()
list_df_mail_asm_columns = df_mail_asm.columns.to_list()
list_df_mail_ss_columns = df_mail_ss.columns.to_list()
list_df_mail_sr_columns = df_mail_sr.columns.to_list()

#designation
sm = "SM"
asm = "ASM"
ss = "SS"
sr = "SR"

end_time = time.time()
print("---------- DONE TRANSFORM! ----------")
execution_time = end_time - start_time
print("Execution time: ", execution_time)
print("-------------------------------------")

#export final sales data
# final_df.to_excel(r'C:\Users\thomas.thai\Downloads\automail\Data sending\Data tổng\final_sales_data_{}.xlsx'.format(max_date))
# print(final_df.info())

# # #sendmail
# send_data_sales_by_number(list_df_mail_rsd_columns, list_of_rsd_numbers, final_df, max_date, sm)
# send_data_sales_by_number(list_df_mail_asm_columns, list_of_asm_numbers, final_df, max_date, asm)
# send_data_sales_by_number(list_df_mail_ss_columns, list_of_ss_numbers, final_df, max_date, ss)
# send_data_sales_by_number(list_df_mail_sr_columns, list_of_sr_numbers, final_df, max_date, sr)

end_time = time.time()  
print("---------- DONE TASK! ----------")
execution_time = end_time - start_time
print("Execution time: ", execution_time)
print("-------------------------------------")

    