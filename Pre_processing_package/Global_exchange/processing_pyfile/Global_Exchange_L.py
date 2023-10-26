import pandas as pd
import numpy as np
import re
import googletrans
from googletrans import Translator
from langdetect import detect
from sklearn.preprocessing import OneHotEncoder
from sklearn.preprocessing import LabelEncoder
import csv
import datetime
from tqdm import tqdm
from tqdm.notebook import tqdm_notebook
from sklearn.preprocessing import MinMaxScaler

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import openpyxl.cell._writer

import os
import openpyxl
import time

'''从这里开始写弹窗'''
def check_columns(df, expected_cols):
    return set(expected_cols).issubset(set(df.columns))

def process_excel(file_path_1, file_path_2):
    data = pd.read_excel(file_path_1)
    expected_cols = ["Program", "Year", "Term", "Status", "Program Date Record: Start Date",
                    "Program Date Record: End Date", "Program Currently Assigned City",
                    "Program Currently Assigned Country", "Program Type", "Student ID",
                    "Country of Citizenship", "Degree Program 1","Degree Program 2",
                    "Points Completed - Total", "Weighted Average", "Postgraduate flag"]
    if not check_columns(data, expected_cols):
        lbl_status.config(text="Please uplode again")
        messagebox.showerror("Error", "Incorrect Globel Exchange Excel form!")
        return
    
    uni_rank = pd.read_csv(file_path_2)
    expected_rank_cols = ['institution', 'Rank']
    if not check_columns(uni_rank, expected_rank_cols):
        lbl_status.config(text="Please uplode again")
        messagebox.showerror("Error", "Incorrect University Rank form!")
        return

    '''从这里开始要处理数据'''
    data = data[expected_cols]
    column_data = data['Term']
    data = data[column_data != "NCP"]
    data['Program Date Record: End Date']  =data['Program Date Record: End Date'].fillna(0)
    data['Program Date RecordFGl: Start Date']  =data['Program Date Record: Start Date'].fillna(0)

    data['Program Date Record: End Date'] = pd.to_datetime(data['Program Date Record: End Date'])
    data['Program Date Record: Start Date'] = pd.to_datetime(data['Program Date Record: Start Date'])
    data['Days'] = (data['Program Date Record: End Date'] - data['Program Date Record: Start Date']).dt.days

    avg_days = {}

    for term in data['Term'].unique():
        filtered_df = data[(data['Term'] == term) & (data['Days'] > 0)]
        avg_days[term] = filtered_df['Days'].mean()

    for term in data['Term'].unique():
        data.loc[(data['Term'] == term) & (data['Days'] <= 0), 'Days'] = avg_days[term]

    data['Days'] = np.round(data['Days']) 

    # Are there program with 0 days?
    #stanardize the name of citizenships
    data['Country of Citizenship']  = data['Country of Citizenship'] .replace({'Country not known': 'NaN',
                                                                        'Not entered': 'NaN','Laos': "Lao People's Democratic Republic",
                                                                        'Republic of Korea': 'Korea, Republic of (South)',
                                                                        'Hong Kong':'Hong Kong (SAR of China)'})

    #assign the status to three types.
    success_substrings = ["Accepted", "Approved", "Nominated", "Finalised", "Committed"]
    unsuccessful_substrings = ["Withdrawn", "Unsuccessful", "Cancelled","Deceased","Reserved","Exemption Requested"]
    pending_substring = ["Pending","Awaiting","Waitlist","Extension"]


    data["Status"] = np.where(data["Status"].str.contains('|'.join(success_substrings)), "Successful", 
                        np.where(data["Status"].str.contains('|'.join(unsuccessful_substrings)), "Unsuccessful", data["Status"]))


    data["Status"] = np.where(data["Status"].str.contains('|'.join(pending_substring)), "Pending", data["Status"])


    today = datetime.date.today()


    for index, row in data.iterrows():
        if any(substring in row['Status'] for substring in pending_substring):
            end_date = pd.to_datetime(row['Program Date Record: End Date'])
            if end_date.date() <= today:
                data.at[index, 'Status'] = "Successful"

    data['Student ID'] = data['Student ID'].fillna(0).astype(int)
    data['Weighted Average'] = data['Weighted Average'].apply(lambda x: re.sub(r'[^\d\.]+', '', str(x)))
    data['Weighted Average'] = pd.to_numeric(data['Weighted Average'], errors='coerce')
    data= data.dropna(subset=['Weighted Average', 'Degree Program 1', 'Degree Program 2','Country of Citizenship'], how='all')  
    data['Postgraduate flag'].replace('False','N',inplace=True)
    data['Postgraduate flag'].replace('True','Y',inplace=True)
    data['Postgraduate flag'].replace('Yes','Y',inplace=True)
    data['Weighted Average'].replace(0, np.nan, inplace=True)
    data['Degree Program 1'].fillna(data['Degree Program 2'], inplace=True)
    data['Degree Program 1'] = data['Degree Program 1'].astype(str)
    data = data.drop(['Degree Program 2'],axis = 1)

    # postgraduate flag consistency
    master_pattern = re.compile(r'\bMaster\b', re.IGNORECASE)
    bachelor_pattern = re.compile(r'\bBachelor\b', re.IGNORECASE)
    postgrad_pattern = re.compile(r'\bPostgraduate\b', re.IGNORECASE)
    grad_pattern = re.compile(r'\bgraduate\b', re.IGNORECASE)
    def update_postgrad_flag(row):
        
        if postgrad_pattern.search(row['Degree Program 1']):
            row['Postgraduate flag'] = 'Y'
        elif master_pattern.search(row['Degree Program 1']):
            row['Postgraduate flag'] = 'Y'
        elif bachelor_pattern.search(row['Degree Program 1']):
            row['Postgraduate flag'] = 'N'
        elif grad_pattern.search(row['Degree Program 1']):
            row['Postgraduate flag'] = 'N'    
        return row


    data = data.apply(update_postgrad_flag, axis=1)

    # weight mean into 4 different options
    remove_n= data.dropna(subset=['Weighted Average'])
    result_year = remove_n.groupby(['Postgraduate flag','Year','Degree Program 1'])['Weighted Average'].mean()
    result_noneYear =  remove_n.groupby(['Postgraduate flag','Degree Program 1'])['Weighted Average'].mean()
    result_OnlyFac =  remove_n.groupby(['Degree Program 1'])['Weighted Average'].mean()
    normal_mean = remove_n['Weighted Average'].mean()
    #means_year = result_year.to_dict()
    means_noneYear = result_noneYear.to_dict()
    means_OnlyFac = result_OnlyFac.to_dict()

    # add average WAM in missing value in three options
    for index, row in data[data['Weighted Average'].isnull()].iterrows():
        flag = row['Postgraduate flag']
        weighted_average = row['Weighted Average']
        points_completed_total = row['Points Completed - Total']

        if flag == 'Y' and (pd.isnull(weighted_average) or weighted_average == 0) and points_completed_total == 0:
            data.at[index, 'Weighted Average'] = 0  #  if flag is 'Y', weighted_average is NaN or 0, and points_completed_total is 0

        year = row['Year']
        school = row['Degree Program 1']


        if (flag, school) in means_noneYear:
            mean = means_noneYear[(flag, school)]
        elif school in means_OnlyFac:
            mean = means_OnlyFac[school]
        else:
            mean = normal_mean  # in this file is 73.75
            
        data.at[index, 'Weighted Average'] = mean

    # wold rank
    rank_table = uni_rank[['institution', 'Rank']]

    # add rank range in rank table
    def get_rank_range(rank):
        if rank <= 10:
            return '1-10'
        elif rank <= 30:
            return '11-30'
        elif rank <= 50:
            return '31-50'
        elif rank <= 100:
            return '51-100'
        elif rank <= 250:
            return '101-250'
        elif rank <= 500:
            return '251-500'
        elif rank <= 1000:
            return '501-1000'
        else:
            return '1001-2000'
            
    rank_table.loc[:, 'rank_range'] = rank_table['Rank'].apply(get_rank_range)
    tqdm.pandas() 
    translator = Translator()

    def safe_translate(text, retries=3):
        for _ in range(retries):
            try:
                return translator.translate(text, dest='en').text
            except Exception as e:
                if _ < retries - 1:  
                    time.sleep(2)
                    continue
                else:   
                    raise e

    def remove_brackets(name):
        name = re.sub(r'\([^()]*\)', '', name)  # remove parentheses and contents
        name = re.sub(r'\s*-\s+.+', '', name)  # remove everything after "-"
        return name.strip()

    # change word to english format and delete and in abbreviation in '()'    
    rank_table.loc[:, 'University'] = rank_table['institution'].progress_apply(lambda x: safe_translate(x) if not detect(x) == 'en' else x)
    rank_table.loc[:, 'University'] = rank_table['University'].progress_apply(remove_brackets)

    # letter learnning for data and rank table
    rank_table.loc[:, 'University'] = rank_table['University'].apply(lambda x: re.sub(r',.*', '', x).strip())
    rank_table.loc[:, 'University'] = rank_table['University'].apply(lambda x: re.sub(r'^[tT]he\s', '', x))
    rank_table.loc[:, 'University'] = rank_table['University'].apply(lambda x: re.sub(r',', '', x))
    data['Program'] = data['Program'].apply(lambda x: re.sub(r',', '', x))

    # vectorized
    university_regex = fr"\b({'|'.join(rank_table['University'].apply(re.escape))})\b"
    data['University'] = data['Program'].str.extract(university_regex, flags=re.IGNORECASE, expand=False).fillna('unknown')
    withRankTable = pd.merge(data,rank_table,on = 'University',how = 'left')
    withRankTable['Rank'] = withRankTable['Rank'].fillna(0).astype(int)
    withRankTable['rank_range'] = withRankTable['rank_range'].fillna('Unknow')
    withRankTable['Year'] = pd.to_datetime(data['Year'],format= '%Y')
    withRankTable['Program Date Record: Start Date'] = pd.to_datetime(data['Program Date Record: Start Date'],format= '%Y')
    withRankTable['Program Date Record: End Date'] = pd.to_datetime(data['Program Date Record: End Date'],format= '%Y')

    # 保存处理后的Excel文件
    output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if output_file:
        withRankTable.to_excel(output_file, index=False)
        lbl_status.config(text="cleaned excel is saved")
        app.after(2000, app.quit)    

def load_file_1():
    global file_path_1
    file_path_1 = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if file_path_1 and file_path_2:
        lbl_status.config(text="Processing...")
        app.after(100, process_excel, file_path_1, file_path_2) 

def load_file_2():
    global file_path_2
    file_path_2 = filedialog.askopenfilename(filetypes=[("Excel files", "*.csv"), ("All files", "*.*")])
    if file_path_1 and file_path_2:
        lbl_status.config(text="Processing...")
        app.after(100, process_excel, file_path_1, file_path_2) 

app = tk.Tk()
app.title("Files Processor")
app.geometry("350x150")

file_path_1 = ""
file_path_2 = ""

load_button_1 = tk.Button(app, text="Load Global Exchange Excel File", command=load_file_1)
load_button_1.pack(pady=10)

load_button_2 = tk.Button(app, text="Load QS Ranking Excel File", command=load_file_2)
load_button_2.pack(pady=10)

lbl_status = tk.Label(app, text="")
lbl_status.pack(pady=10)

app.mainloop()