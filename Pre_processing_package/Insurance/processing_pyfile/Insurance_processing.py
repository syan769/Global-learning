import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import openpyxl.cell._writer

def check_columns(df, expected_cols):
    return set(expected_cols).issubset(set(df.columns))

def process_excel(input_file):
    insurance = pd.read_excel(input_file)
    expected_cols = ['Year', 'Itinerary','Student ID','Purpose of Travel','Does your complete journey exceed 365 days?','Level of Study','Faculty','Date of Departure','Date of Return']
    if not check_columns(insurance, expected_cols):
        lbl_status.config(text="Please uplode again")
        messagebox.showerror("Error", "Incorrect Excel form!")
        return
    
    # 这里开始是表格处理
    insurance = insurance[['Year', 'Itinerary','Student ID','Purpose of Travel','Does your complete journey exceed 365 days?','Level of Study','Faculty','Date of Departure','Date of Return']]
    # 把ltinerary中的string拆开
    insurance[['city','country','continent', 'date_range']] = insurance['Itinerary'].str.extract(r'^(.*), (.*?) \((.*?)\) (Start Date: .* - End Date: .*)$')
    # 将 "Start Date" 和 "End Date" 列转换为日期类型
    insurance = insurance.dropna(subset=['continent', 'country', 'city', 'date_range'])
    insurance[['Start Date', 'End Date']] = insurance['date_range'].str.extract(r'^Start Date: (.*) - End Date: (.*)$')
    insurance['Start Date'] = pd.to_datetime(insurance['Start Date'], format='%d/%m/%Y')
    insurance['End Date'] = pd.to_datetime(insurance['End Date'], format='%d/%m/%Y')
    insurance = insurance.drop(['Itinerary','date_range'], axis = 1)

    # 把student ID变成int和nan值变0
    insurance['Student ID'] = insurance['Student ID'].fillna(0)
    insurance['Student ID'] = insurance['Student ID'].astype('int64')
    # 将重复的行都删掉，assume重复计算
    insurance = insurance.drop_duplicates()
    # 把date of departure和return变成duration，保留date of departure
    insurance['Duration'] = insurance['Date of Return'] - insurance['Date of Departure']
    insurance['Program_duration'] = insurance['End Date'] - insurance['Start Date']
    # 如果Duration<Program_duration,就将Program_duration填入Duration
    insurance['Duration'] = insurance.apply(lambda x: x['Program_duration'] if x['Duration'] <= x['Program_duration'] else x['Duration'], axis=1)
    # 将Duration变成分类模式，分成<14两周,14-30一个月内,30-90一个季度内，90-180半年内，180- 一年或以上
    insurance['Duration'] = insurance['Duration'].dt.days.astype(int)

    # 将天数分成不同的类别
    bins = [-1, 14, 30, 90, 180, insurance['Duration'].max()]
    labels = ['2weeks', '1month', '3months', 'half_year', 'one year or more']
    insurance['Duration_Category'] = pd.cut(insurance['Duration'], bins=bins, labels=labels)
    # 年份变成datetime格式
    insurance['Year'] = pd.to_datetime(insurance['Year'],format='%Y')
    insurance = insurance.drop(['Date of Return','Start Date','End Date','Program_duration'],axis=1)
    
    # 保存处理后的Excel文件
    output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if output_file:
        insurance.to_excel(output_file, index=False)
        lbl_status.config(text="cleaned excel is saved")
        app.after(2000, app.quit)  # 2秒后关闭窗口

def load_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    # UI刷新显示"Processing..."
    if file_path:
        lbl_status.config(text="Processing...")
        app.after(100, process_excel, file_path)  

app = tk.Tk()
app.title("Insurance Processor")
app.geometry("300x100")

load_button = tk.Button(app, text="Load Excel File", command=load_file)
load_button.pack(pady=20)

lbl_status = tk.Label(app, text="")
lbl_status.pack(pady=20)

app.mainloop()
