########Import Commands##########
import pandas as pd
import openpyxl as xl
import numpy as np


#######Loading in Workbook, isolating sheet formats########
wb = xl.load_workbook('Fra_Har_Wan_dataset_11_18.xlsm')
turbine = wb.sheetnames[2:]


maindf = pd.DataFrame(['Final Cleaned Dataset'])
with pd.ExcelWriter('Fra_Har_Wan_testing.xlsx') as writer:
    maindf.to_excel(writer, sheet_name='main')


########Row Deletion and Statistics Collection###########
#Fra, Har, Wan Sheets
for sheet in turbine:
    data = wb[sheet]
    old_df = pd.DataFrame(data.values)
    df = old_df[[0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31]]
    unit_desc = df.iloc[[1,3],:]
    top_col_data = df.iloc[1:9,:]
    headers = df.iloc[0,:]
    lst_headers = list(headers)
    unit = df.iloc[3,:]
    df = df.iloc[9:,:]
    temp_cols=[]
    vib_cols=[]
    volt_cols=[]
    kicked_rows = []
    for (col_name, col_data) in unit.iteritems():
        if col_data.upper() == 'DEG F':
            temp_cols.append(col_name)
        if col_data.upper() == 'MILS':
            vib_cols.append(col_name)
        if col_data.upper() == 'VOLTS':
            volt_cols.append(col_name)
        if col_data.upper() == 'ENGUNITS':
            date_cols = col_name
    kick_len_volt = len(volt_cols)
    kick_len_vib = len(vib_cols)
    for (row_name, row_data) in df.iterrows():
        if row_data[volt_cols].isna().sum() == kick_len_volt:
            kicked_rows.append(row_name)
        elif row_data[vib_cols].isna().sum() == kick_len_vib:
            kicked_rows.append(row_name)

    print('Sheet Name: ' + sheet + ' , Missing Rows: ' + str(len(kicked_rows)))
    df = df.drop(kicked_rows)
    for (col_name, col_data) in df.iteritems():
        if col_name in temp_cols:
            print('Temp Col' + str(col_name) + ' ' + str(df[col_name].isna().sum()))
        elif col_name in vib_cols:
            print('Vib Col' + str(col_name) + ' ' + str(df[col_name].isna().sum()))
        elif col_name in volt_cols:
            print('Volt Col' + str(col_name) + ' ' + str(df[col_name].isna().sum()))
        else:
            print('Other Col' + str(col_name) + ' ' + str(df[col_name].isna().sum()))
    num_both_missing = []
    brg_temp_col = [4, 5, 6, 7, 8, 9, 10, 11, 12, 13]
    for i in brg_temp_col:
        df[i] = df[i].mask(((df[i] < 80) | (df[i] > 280)), np.nan)

    vib_missing=[]
    for i in vib_cols:
        df[i] = df[i].mask((df[i] > 12), np.nan)
        vib_missing.append(df[i].isna().sum())

    vol_missing = []
    for i in volt_cols:
        df[i] = df[i].mask(((df[i] > -6) | (df[i] < -18)), np.nan)
        vol_missing.append(df[i].isna().sum())


    for i in range(0, len(brg_temp_col), 2):
        new_col_name = "Max" + headers.iloc[brg_temp_col[i]]
        lst_headers.append(new_col_name)
        df[new_col_name] = df[[brg_temp_col[i], brg_temp_col[i+1]]].max(axis=1)
        num_both_missing.append(df[new_col_name].isna().sum())
        top_col_data[new_col_name] = np.nan

    lst_headers[0] = 'date'
    df.columns = lst_headers
    top_col_data.columns = lst_headers
    print(num_both_missing)
    print(vib_missing)
    print(vol_missing)
    print(df.head())

    frames = [top_col_data, df]
    finaldf = pd.concat(frames)
    with pd.ExcelWriter('Fra_Har_Wan_testing.xlsx', mode='a') as writer:
        finaldf.to_excel(writer, sheet_name= sheet)


# #Row Ct 1,2,3 sheets
# for sheet in row_ct_123:
#     data = wb[sheet]
#     old_df = pd.DataFrame(data.values)
#     df = old_df[
#         [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26]]
#     unit_desc = df.iloc[[1, 3], :]
#     unit = df.iloc[3, :]
#     top_col_data = df.iloc[1:9, :]
#     headers = df.iloc[0,:]
#     lst_headers = list(headers)
#     df = df.iloc[9:, :]
#     temp_cols = []
#     vib_cols = []
#     volt_cols = []
#     kicked_rows = []
#     for (col_name, col_data) in unit.iteritems():
#         if col_data.upper() == 'DEG F':
#             temp_cols.append(col_name)
#         if col_data.upper() == 'IN/SEC':
#             vib_cols.append(col_name)
#         if col_data.upper() == 'VOLTS':
#             volt_cols.append(col_name)
#         if col_data.upper() == 'ENGUNITS':
#             date_cols = col_name
#     kick_len_volt = len(volt_cols)
#     kick_len_vib = len(vib_cols)
#     for (row_name, row_data) in df.iterrows():
#         if row_data[volt_cols].isna().sum() == kick_len_volt:
#             kicked_rows.append(row_name)
#         elif row_data[vib_cols].isna().sum() == kick_len_vib:
#             kicked_rows.append(row_name)
#
#     print('Sheet Name: ' + sheet + ' , Missing Rows: ' + str(len(kicked_rows)))
#     df = df.drop(kicked_rows)
#     for (col_name, col_data) in df.iteritems():
#         if col_name in temp_cols:
#             print('Temp Col' + str(col_name) + ' ' + str(df[col_name].isna().sum()))
#         elif col_name in vib_cols:
#             print('Vib Col' + str(col_name) + ' ' + str(df[col_name].isna().sum()))
#         elif col_name in volt_cols:
#             print('Volt Col' + str(col_name) + ' ' + str(df[col_name].isna().sum()))
#         else:
#             print('Other Col' + str(col_name) + ' ' + str(df[col_name].isna().sum()))
#
#     num_both_missing = []
#     brg_temp_col = [4, 5, 6, 7, 8, 9, 10, 11]
#     for i in brg_temp_col:
#         df[i] = df[i].mask(((df[i] < 80) | (df[i] > 280)), np.nan)
#
#     vib_missing = []
#     for i in vib_cols:
#         df[i] = df[i].mask((df[i] > 12), np.nan)
#         vib_missing.append(df[i].isna().sum())
#
#     vol_missing = []
#     for i in volt_cols:
#         df[i] = df[i].mask(((df[i] > -6) | (df[i] < -18)), np.nan)
#         vol_missing.append(df[i].isna().sum())
#
#     for i in range(0, len(brg_temp_col), 2):
#         new_col_name = "Max" + headers.iloc[brg_temp_col[i]]
#         lst_headers.append(new_col_name)
#         df[new_col_name] = df[[brg_temp_col[i], brg_temp_col[i+1]]].max(axis=1)
#         num_both_missing.append(df[new_col_name].isna().sum())
#         top_col_data[new_col_name] = np.nan
#
#     lst_headers[0] = 'date'
#     df.columns = lst_headers
#     top_col_data.columns = lst_headers
#     print(num_both_missing)
#     print(vib_missing)
#     print(vol_missing)
#     print(df.head())
#
#     frames = [top_col_data, df]
#     finaldf = pd.concat(frames)
#     with pd.ExcelWriter('final_cleaned_dataset.xlsx', mode='a') as writer:
#         finaldf.to_excel(writer, sheet_name= sheet)

# #Row Ct 4,5 sheets
# for sheet in row_ct_45:
#     data = wb[sheet]
#     old_df = pd.DataFrame(data.values)
#     df = old_df[
#         [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26,27]]
#     unit_desc = df.iloc[[1, 3], :]
#     unit = df.iloc[3, :]
#     top_col_data = df.iloc[1:9, :]
#     headers = df.iloc[0,:]
#     lst_headers = list(headers)
#     df = df.iloc[9:, :]
#     temp_cols = []
#     vib_cols = []
#     volt_cols = []
#     kicked_rows = []
#     for (col_name, col_data) in unit.iteritems():
#         if col_data.upper() == 'DEG F':
#             temp_cols.append(col_name)
#         if col_data.upper() == 'IN/SEC':
#             vib_cols.append(col_name)
#         if col_data.upper() == 'VOLTS':
#             volt_cols.append(col_name)
#         if col_data.upper() == 'ENGUNITS':
#             date_cols = col_name
#     kick_len_volt = len(volt_cols)
#     kick_len_vib = len(vib_cols)
#     for (row_name, row_data) in df.iterrows():
#         if row_data[volt_cols].isna().sum() == kick_len_volt:
#             kicked_rows.append(row_name)
#         elif row_data[vib_cols].isna().sum() == kick_len_vib:
#             kicked_rows.append(row_name)
#
#     print('Sheet Name: ' + sheet + ' , Missing Rows: ' + str(len(kicked_rows)))
#     df = df.drop(kicked_rows)
#     for (col_name, col_data) in df.iteritems():
#         if col_name in temp_cols:
#             print('Temp Col' + str(col_name) + ' ' + str(df[col_name].isna().sum()))
#         elif col_name in vib_cols:
#             print('Vib Col' + str(col_name) + ' ' + str(df[col_name].isna().sum()))
#         elif col_name in volt_cols:
#             print('Volt Col' + str(col_name) + ' ' + str(df[col_name].isna().sum()))
#         else:
#             print('Other Col' + str(col_name) + ' ' + str(df[col_name].isna().sum()))
#
#     num_both_missing = []
#     brg_temp_col = [5, 6, 7, 8, 9, 10, 11, 12]
#     for i in brg_temp_col:
#         df[i] = df[i].mask(((df[i] < 80) | (df[i] > 280)), np.nan)
#
#     vib_missing = []
#     for i in vib_cols:
#         df[i] = df[i].mask((df[i] > 12), np.nan)
#         vib_missing.append(df[i].isna().sum())
#
#     vol_missing = []
#     for i in volt_cols:
#         df[i] = df[i].mask(((df[i] > -6) | (df[i] < -18)), np.nan)
#         vol_missing.append(df[i].isna().sum())
#
#     for i in range(0, len(brg_temp_col), 2):
#         new_col_name = "Max" + headers.iloc[brg_temp_col[i]]
#         lst_headers.append(new_col_name)
#         df[new_col_name] = df[[brg_temp_col[i], brg_temp_col[i+1]]].max(axis=1)
#         num_both_missing.append(df[new_col_name].isna().sum())
#         top_col_data[new_col_name] = np.nan
#
#     lst_headers[0] = 'date'
#     df.columns = lst_headers
#     top_col_data.columns = lst_headers
#     print(num_both_missing)
#     print(vib_missing)
#     print(vol_missing)
#     print(df.head())
#
#     frames = [top_col_data, df]
#     finaldf = pd.concat(frames)
#     with pd.ExcelWriter('final_cleaned_dataset.xlsx', mode='a') as writer:
#         finaldf.to_excel(writer, sheet_name= sheet)

# #Addison sheets
# for sheet in addison_sheets:
#     data = wb[sheet]
#     old_df = pd.DataFrame(data.values)
#     df = old_df[
#         [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]]
#     unit_desc = df.iloc[[1, 3], :]
#     unit = df.iloc[3, :]
#     top_col_data = df.iloc[1:9, :]
#     headers = df.iloc[0,:]
#     lst_headers = list(headers)
#     df = df.iloc[9:, :]
#     temp_cols = []
#     vib_cols = []
#     volt_cols = []
#     kicked_rows = []
#     for (col_name, col_data) in unit.iteritems():
#         if col_data.upper() == 'DEG F':
#             temp_cols.append(col_name)
#         if col_data.upper() == 'IN/S':
#             vib_cols.append(col_name)
#         if col_data.upper() == 'ENGUNITS':
#             date_cols = col_name
#     kick_len_vib
#     for (row_name, row_data) in df.iterrows():
#         if row_data[vib_cols].isna().sum() == kick_len_vib:
#             kicked_rows.append(row_name)
#
#     print('Sheet Name: ' + sheet + ' , Missing Rows: ' + str(len(kicked_rows)))
#     df = df.drop(kicked_rows)
#     for (col_name, col_data) in df.iteritems():
#         if col_name in temp_cols:
#             print('Temp Col' + str(col_name) + ' ' + str(df[col_name].isna().sum()))
#         elif col_name in vib_cols:
#             print('Vib Col' + str(col_name) + ' ' + str(df[col_name].isna().sum()))
#         else:
#             print('Other Col' + str(col_name) + ' ' + str(df[col_name].isna().sum()))
#
#     num_both_missing = []
#     brg_temp_col = [4, 5, 6, 7, 8, 9, 10, 11]
#     for i in brg_temp_col:
#         df[i] = df[i].mask(((df[i] < 80) | (df[i] > 280)), np.nan)
#
#     vib_missing = []
#     for i in vib_cols:
#         df[i] = df[i].mask((df[i] > 12), np.nan)
#         vib_missing.append(df[i].isna().sum())
#
#     vol_missing = []
#     for i in volt_cols:
#         df[i] = df[i].mask(((df[i] > -6) | (df[i] < -18)), np.nan)
#         vol_missing.append(df[i].isna().sum())
#
#     for i in range(0, len(brg_temp_col), 2):
#         new_col_name = "Max" + headers.iloc[brg_temp_col[i]]
#         lst_headers.append(new_col_name)
#         df[new_col_name] = df[[brg_temp_col[i], brg_temp_col[i+1]]].max(axis=1)
#         num_both_missing.append(df[new_col_name].isna().sum())
#         top_col_data[new_col_name] = np.nan
#
#     lst_headers[0] = 'date'
#     df.columns = lst_headers
#     top_col_data.columns = lst_headers
#     print(num_both_missing)
#     print(vib_missing)
#     print(vol_missing)
#     print(df.head())
#
#     frames = [top_col_data, df]
#     finaldf = pd.concat(frames)
#     with pd.ExcelWriter('final_cleaned_dataset.xlsx', mode='a') as writer:
#         finaldf.to_excel(writer, sheet_name= sheet)

#######Imputation Process#########