"""
Simple app to compare two excel sheets basing on key column.
First file sholuld old one and the second new one.
Changes are marked in green(new data), yellow and red(missing data)
In addition values that changed(yellow) contains old and new value "OLD -> NEW"
Result excel is saves as new file in location on second(new) file. 
"""


import pandas as pd
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox as msb
import os

def compare_dataframes2(df1, df2, key_column):
    if not key_column in df1.columns and not key_column in df2.columns:
        #list of columns that are in both df1 and df2 
        common_columns = df1.columns.intersection(df2.columns).tolist()

        #list of columns that are in df1 but not in df2
    #check if key column in both dataframes has duplicated values
    df1[key_column] = df1[key_column].astype(str)
    df2[key_column] = df2[key_column].astype(str)

    if df1[key_column].duplicated().any() or df2[key_column].duplicated().any():
        #list of duplicated keys in key columns (df1 and df2)
        duplicated_keys = df1[key_column][df1[key_column].duplicated()].tolist()
        duplicated_keys.extend(df2[key_column][df2[key_column].duplicated()].tolist())
        
        #raise Exception('Key column has duplicated values')
        msb.showerror('Error', 'Key column has duplicated values ' + str(duplicated_keys))
        return

    #new rows in df2
    new_rows = df2[~df2[key_column].isin(df1[key_column])][key_column].tolist()

    # Create a copy of df2 to hold the modified data
    df2_modified = df2.copy()

    missing_columns = df1.columns.difference(df2.columns).tolist()
    missing_column_df1 = df2.columns.difference(df1.columns).tolist()
    print('missing_columns df1', missing_column_df1)
    
    new_missing_columns = [col + '_MISSING' for col in missing_columns]
    #columns, that are in df2 but not in df1
    new_cols = df2.columns.difference(df1.columns).tolist()

    ##############################################################################################
    # Find rows in df1 that are missing from df2
    missing_mask = ~df1[key_column].isin(df2[key_column])
    missing_rows_key = df1[missing_mask][key_column].tolist()
    
    mask = df1[key_column].isin(missing_rows_key)
    df1.loc[mask,key_column] = df1.loc[mask,key_column]+ '_MISSING'
    missing_rows_key = df1[missing_mask][key_column].tolist()

    missing_rows = df1[~df1[key_column].isin(df2[key_column])]

    df2_modified.fillna('', inplace=True)
    df1.fillna('', inplace=True)


    df2_modified = pd.concat([df2_modified,missing_rows], ignore_index=True, sort=False)
    df2_modified.reset_index(drop=True, inplace=True)

    df2_modified.drop(missing_columns, axis=1, inplace=True)
    ############
    keys  =  df2_modified[key_column].tolist()
    keys_dict = {}
    for i, key in enumerate(keys):
        keys_dict[key] = i
    ########3###

    keys_df1  =  df1[key_column].tolist()
    keys_dict_1 = {}
    for i, key in enumerate(keys_df1):
        keys_dict_1[key] = i
    ############
    cols = df2_modified.columns.tolist()
    cols_dict = {}
    for i, col in enumerate(cols):
        cols_dict[col] = i
    ############
    for col in missing_columns:
        new_col = col + '_MISSING'  
        df2_modified[new_col] = ''
        for key,val in keys_dict_1.items():
            #print(key, val)
            df2_modified.loc[keys_dict[key],new_col] = df1.loc[val,col]
    
    def cc(x):
        color = 'background-color: gold'
        dfx = pd.DataFrame('', index=x.index, columns=x.columns)
        for col in cols_dict.keys():
            if col in ['index', key_column] + missing_column_df1: continue
            if not col in df1.columns: continue
            df1[col] = df1[col].astype(str)
            df2_modified[col] = df2_modified[col].astype(str)
            for key,val in keys_dict.items():
                if key in missing_rows_key + new_rows: continue
                if df1.loc[val, col] != df2_modified.loc[val, col]:
                    df2_modified.loc[val, col] = df1.loc[val, col] + ' -> ' + df2_modified.loc[val, col]
                    dfx.loc[val, col] = color 
                    dfx.loc[val, key_column] = color
        return dfx

    def missing_color_cells(val):
        if str(val).endswith('_MISSING'):
            return 'background-color:tomato'

    def red_cols(s):
        color = 'tomato'
        return 'background-color: %s' % color

    def green_cols(s):
        color = 'lightgreen'
        return 'background-color: %s' % color

    def style_specific_cell(x):
        color = 'background-color: lightgreen'
        dfx = pd.DataFrame('', index=x.index, columns=x.columns)
        for r in new_rows:
            dfx.iloc[keys_dict[r], 0] = color
        return dfx

    df2_modified.reset_index(drop=False, inplace=True)

    for col in ['index','level_0']:
        if col in df2_modified.columns:
            df2_modified.drop(col, axis=1, inplace=True)

    df2_modified_styled = df2_modified.style.apply(cc, axis=None)#.hide_index()
    df2_modified_styled.applymap(red_cols, subset=new_missing_columns)
    df2_modified_styled.applymap(missing_color_cells)
    df2_modified_styled.applymap(green_cols, subset=new_cols)
    df2_modified_styled.apply(style_specific_cell, axis=None)

    df2_modified_styled.set_properties(**{'text-align': 'center'})
    #df2.set_index('KEY', inplace=True)
    #df2_modified.fillna('', inplace=True)

    return df2_modified_styled

def select_files_and_key_column(initialdir = '', priority_columns_list = []):
    import json
    import datetime

    def compare_files(temp_file1, temp_file2, tab1, tab2, key_column):
        compare_data = {'file1':temp_file1, 'file2':temp_file2, 'tab':tab1, 'key_column':key_column}
        df1 = pd.read_excel(temp_file1, sheet_name=tab1)
        df2 = pd.read_excel(temp_file2, sheet_name=tab2)

        df_comp = compare_dataframes2(df1,df2, key_column = key_column) 
        path = os.path.dirname(temp_file2) + '/Compare/'
        if not os.path.exists(path):
            os.makedirs(path)
        date_stamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        file_name = path + os.path.basename(temp_file2).split('.')[0] + '___' + tab1 + '___' + date_stamp + '.xlsx'
        #save compare_data to file
        with open(file_name[:-5] + '.json', 'w') as fp:
            json.dump(compare_data, fp)
        df_comp.to_excel(file_name, index=False)
        os.startfile(file_name)

    if len(priority_columns_list) == 0:
        priority_columns_list = ['SP_NUMBER','ID','RefNo','RefNo2','frameRefNo','unitRefNo' ]
    if len(initialdir) == 0:
        initialdir = "C:/"

    temp_file1 = filedialog.askopenfilename(initialdir = initialdir, title = "PREVIOUS / OLD file", filetype = (("excel files","*.xlsx"),("all files","*.*")))
    temp_file2 = filedialog.askopenfilename(initialdir = initialdir, title = "CURRENT / NEW file", filetype = (("excel files","*.xlsx"),("all files","*.*")))

    #check if file1 is older than file2
    if os.path.getmtime(temp_file1) > os.path.getmtime(temp_file2):
        msb.showwarning("Warning", "File 1 is newer than file 2. Please select files in correct order")
        ask = msb.askokcancel("Warning", "File 1: " + temp_file1 + "\nFile 2: " + temp_file2 + "\nFile 1 is newer than file 2. Want to continue?")
        if ask == False:
            return

    if temp_file1 == '' or temp_file2 == '':
        msb.showerror("Error", "Please select input files")
        return

    #create popup window to select tabs in worksheet an key column
    popup_comp = Toplevel()
    popup_comp.geometry('400x250')
    popup_comp.grab_set()
    popup_comp.title("Compare input data files ") 
    popup_comp.resizable(False, False)

    def update_columns(*args):
        tab_name = combo_tab1.get()
        df = pd.read_excel(temp_file1, sheet_name=tab_name)
        column_names = df.columns.tolist()
        #move priority columns to the beginning of the list
        for col in priority_columns_list:
            if col in column_names:
                column_names.remove(col)
                column_names.insert(0, col)
        combo_key_column['values'] = column_names

    #check if number of tabs in excel file is more than 1
    if len(pd.ExcelFile(temp_file1).sheet_names) > 1:
        #create label and combobox to select tab in worksheet 
        #when combobox is changed, update list of columns in combobox for key column        
        label_tab1 = Label(popup_comp, text="Select tab in 1st file")
        label_tab1.grid(row=0, column=0, padx=10, pady=10)
        combo_tab1 = ttk.Combobox(popup_comp, state="readonly")
        combo_tab1.bind("<<ComboboxSelected>>", update_columns)
        combo_tab1['values'] = pd.ExcelFile(temp_file1).sheet_names
        combo_tab1.grid(row=0, column=1, padx=10, pady=10)
        combo_tab1.current(0)
    else:
        combo_tab1 = pd.ExcelFile(temp_file1).sheet_names[0]

    
    if len(pd.ExcelFile(temp_file2).sheet_names) > 1:
        #create label and combobox to select tab in worksheet
        label_tab2 = Label(popup_comp, text="Select tab in 2nd file")
        label_tab2.grid(row=1, column=0, padx=10, pady=10)
        combo_tab2 = ttk.Combobox(popup_comp, state="readonly")
        combo_tab2['values'] = pd.ExcelFile(temp_file2).sheet_names
        combo_tab2.grid(row=1, column=1, padx=10, pady=10)
        combo_tab2.current(0)
    else:
        combo_tab2 = pd.ExcelFile(temp_file2).sheet_names[0]

    #create label and combobox to select key column in first file and selected tab (if more than 1 tab)
    columns = pd.read_excel(temp_file1).columns.tolist()
    #move priority columns to the beginning of the list
    for col in priority_columns_list:
        if col in columns:
            columns.remove(col)
            columns.insert(0, col)
    label_key_column = Label(popup_comp, text="Select key column in 1st file")
    label_key_column.grid(row=10, column=0, padx=10, pady=10)
    combo_key_column = ttk.Combobox(popup_comp, state="readonly")
    combo_key_column['values'] = list(pd.read_excel(temp_file1).columns)
    combo_key_column.grid(row=10, column=1, padx=10, pady=10)
    combo_key_column.current(0)

    #create button to start comparison
    button_compare = Button(popup_comp, text="Compare", command=lambda: compare_files(temp_file1, temp_file2, combo_tab1.get(), combo_tab2.get(), combo_key_column.get()))
    button_compare.grid(row=22, column=0, padx=10, pady=10)
    button_cancel = Button(popup_comp, text="Cancel", command=popup_comp.destroy)
    button_cancel.grid(row=22, column=1, padx=10, pady=10)

    popup_comp.mainloop()


def test_compare():

    # create the dataframe
    df1 = pd.DataFrame(values1, columns=['KEY', 'C2', 'C3', 'C4'])
    df2 = pd.DataFrame(values2, columns=['KEY', 'C2', 'C3', 'C5'])

    df1.reset_index(drop=True, inplace=True)
    df2.reset_index(drop=True, inplace=True)

    df_comp = compare_dataframes2(df1,df2,   key_column = 'SP_NUMBER')  
    df_comp.to_excel(initialdir + 'dataframe.xlsx', index=False)
    os.startfile(initialdir + 'dataframe.xlsx')


def compare_2_excels():
    initialdir = "M:/00 MAS_TASK/6.Output/"
    priority_columns_list = ['SP_NUMBER','ID','RefNo','RefNo2','frameRefNo','unitRefNo' ]

    select_files_and_key_column(initialdir, priority_columns_list)


if __name__ == "__main__": 
    compare_2_excels()

