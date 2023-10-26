import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import openpyxl.cell._writer

log_book = []
log_content = []

def check_columns(df, expected_cols):
    return set(expected_cols).issubset(set(df.columns))

def process_excel(file_path_1, file_path_2):
    raw_citizenship = pd.read_excel(file_path_1)
    raw_worldcities = pd.read_excel(file_path_2)
    expected_cols = ['Year', 'Program', 'Placement Location - City', 'Placement Location - State', 'Placement Location - Country', 'Placement Start Date', 'Placement End Date', "Country of Citizenship", 'ID#']
    if not check_columns(raw_citizenship, expected_cols):
        lbl_status.config(text="Please upload again")  # Small typo fixed: uplode -> upload
        messagebox.showerror("Error", "Incorrect Excel form!")
        return
    ''' 这里开始是表格处理''' 
    # Pre-processing for the dataset "faculty.xlsx"
    # Note: Please makesure the file name called "faculty.xlsx" to avoid a reading error
    # Step 1
    # read excel file in python and check header
    # Step 2:
    # Creat a log book to store the problem records
    def add_entry(log_book, entry_id, city_b, country_b, city_a, country_a):
        entry = {
            "ID": entry_id,
            "City_before": city_b,
            "Country_before": country_b,
            "City_after": city_a,
            "Country_after": country_a
        }
        log_book.append(entry)
        
    def display_entries(log_book):
        log_content = []  
        for entry in log_book:
            log_content.append("ID: " + str(entry["ID"]))
            log_content.append("City Before: " + entry["City_before"])
            log_content.append("Country Before: " + entry["Country_before"])
            log_content.append("City after: " + entry["City_after"])
            log_content.append("Country after: " + entry["Country_after"])
            log_content.append("-" * 20)  # Separator between entries

        return "\n".join(log_content)

    # Step 3.1: 
    # Check empty columns
    list_city_nan = raw_citizenship['Placement Location - City'][pd.isna(raw_citizenship['Placement Location - City'])].index.tolist()

    # Step 4
    # Creat a standarized city-country dictionary called  to correct error
    dict_worldcities = {}
    new_worldcities = raw_worldcities[['city_ascii', 'country']]

    for col_name, col_data in new_worldcities.iterrows():
        
        if col_data['country'] in dict_worldcities:
            dict_worldcities[col_data['country']].append(col_data['city_ascii'])
        else:
            dict_worldcities[col_data['country']] = [col_data['city_ascii']]


    raw_citizenship = raw_citizenship.dropna(subset=['Placement Location - City'])


    # Looking for the problem country
    problem_country = []
    problem_row = []
    for col_name, col_data in raw_citizenship.iterrows():
        country = col_data['Placement Location - Country']
        if country not in dict_worldcities:
            ID = col_data['ID#']
            if country == 'United States Of America' or country == 'USA':
                normalized_country = 'United States'
            elif country == 'The Netherlands':
                normalized_country = 'Netherlands'
            elif country == 'Kathmandu':
                normalized_country = 'Nepal'
            
            raw_citizenship.loc[(raw_citizenship['ID#'] == ID), 'Placement Location - Country'] = normalized_country
            add_entry(log_book, ID, col_data['Placement Location - City'], country, col_data['Placement Location - City'], normalized_country)

    log_content = display_entries(log_book)


    problem_city = []
    problem_city_row = []
    for col_name, col_data in raw_citizenship.iterrows():

        if col_data['Placement Location - City'] not in dict_worldcities[col_data['Placement Location - Country']]:
            
            problem_city.append(col_data['Placement Location - City'])
            problem_city_row.append(col_data)


    # remove 'city', 'City', 'South' append on the city
    raw_citizenship['Placement Location - City'] = raw_citizenship['Placement Location - City'].str.replace(' City', '')
    raw_citizenship['Placement Location - City'] = raw_citizenship['Placement Location - City'].str.replace(' city', '')
    raw_citizenship['Placement Location - City'] = raw_citizenship['Placement Location - City'].str.replace('South ', '')

    # New version algorithm
    # To improve matching scores, we match the city name with the most frequent appeared in the dataset.

    # 1. build a dictionary to store all the vaild country-city pair appeared in the dataset
    # 2. According to the country name, match the simialr city with error city
    # 3. print out top 3 similar city name and the matching scores
    # 4. check validation
    high_freq_dic = {}
    for col_name, col_data in raw_citizenship.iterrows():
        country = col_data['Placement Location - Country']
        city = col_data['Placement Location - City']
        if (country in dict_worldcities) and (city in dict_worldcities[country]):
            if country in high_freq_dic:
                if city not in high_freq_dic[country]:
                    high_freq_dic[country].append(city)
            else:
                high_freq_dic[country] = [city]


    # Match algorithm
    def match_city_to_country(city_name, country, dict_worldcities):
        best_match = None
        highest_similarity = 0

        for dic_city in dict_worldcities[country]:
            similarity = calculate_similarity(city_name, dic_city)
            if similarity > highest_similarity:
                highest_similarity = similarity
                best_match = (city_name, dic_city)

        return best_match

    def calculate_similarity(str1, str2):
        set1 = set(str1)
        set2 = set(str2)
        
        intersection = len(set1 & set2)
        union = len(set1 | set2)
        
        similarity = intersection / union
        return similarity


    # Test match performance
    # city_name = "Helbournne"
    # Using high freq dic instead of standard world city dictionary to improve the performance of algorithm
    # best_match = match_city_to_country(city_name, "Australia", high_freq_dic)
    # print(best_match)

    # -----------------------------------------------------------------------------------------------------------------------


    # Earliest version to processing human error on city which
    citizenship_after_drop = raw_citizenship
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - State'] == 'Tasmania') & (citizenship_after_drop['Placement Location - City'] == 'Brighton'), 'Placement Location - City'] = 'Hobart'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - State'] == 'Manama') & (citizenship_after_drop['Placement Location - City'] == 'Manama'), 'Placement Location - Country'] = 'Bahrain'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - State'] == 'New South Wales') & (citizenship_after_drop['Placement Location - City'] == 'Wollogorang'), 'Placement Location - City'] = 'Wollongong'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - State'] == 'Queensland') & (citizenship_after_drop['Placement Location - City'] == 'Trinity Beach'), 'Placement Location - City'] = 'Cairns'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - State'] == 'New South Wales') & (citizenship_after_drop['Placement Location - City'] == 'Moorwatha'), 'Placement Location - City'] = 'Albury'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - State'] == 'Queensland') & (citizenship_after_drop['Placement Location - City'] == 'Carrara'), 'Placement Location - City'] = 'Gold Coast'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - State'] == 'Northern Territory') & (citizenship_after_drop['Placement Location - City'] == 'Tennant Creek'), 'Placement Location - City'] = 'Darwin'

    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'singapore'), 'Placement Location - City'] = 'Singapore'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Ho Chi Minh'), 'Placement Location - City'] = 'Ho Chi Minh City'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Winchester'), 'Placement Location - City'] = 'Ottawa'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Fort McMurray'), 'Placement Location - City'] = 'Edmonton'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Matredal'), 'Placement Location - City'] = 'Bergen'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Greater Sudbury'), 'Placement Location - City'] = 'Sudbury'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - Country'] == 'Malaysia'), 'Placement Location - City'] = 'Kuala Lumpur'

    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'FOOTSCRAY'), 'Placement Location - City'] = 'Melbourne'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - State'] == 'Nusa Tenggara Barat'), 'Placement Location - City'] = 'Lembok'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - State'] == 'Bali'), 'Placement Location - City'] = 'Denpasar'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - Country'] == 'Denmark'), 'Placement Location - City'] = 'Copenhagen'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - Country'] == 'Ghana'), 'Placement Location - City'] = 'Accra'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - Country'] == 'New Caledonia'), 'Placement Location - City'] = 'Noumea'

    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Oak Bluff'), 'Placement Location - City'] = 'Winnipeg'

    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Yeoncheon'), 'Placement Location - City'] = 'Pocheon'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Uiwang-si'), 'Placement Location - City'] = 'Anyang'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'vancouver'), 'Placement Location - City'] = 'Vancouver'

    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Niarobi'), 'Placement Location - City'] = 'Nairobi'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Den Haag'), 'Placement Location - City'] = 'The Hague'

    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Den Haag'), 'Placement Location - City'] = 'The Hague'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Chapinero'), 'Placement Location - City'] = 'Bogota'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Tartu Maarkon'), 'Placement Location - City'] = 'Tartu'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Langley Township'), 'Placement Location - City'] = 'Langley'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Baluwatar'), 'Placement Location - City'] = 'Kathmandu'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Mexico'), 'Placement Location - City'] = 'Mexico City'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Manu learning centre'), 'Placement Location - City'] = 'Cusco'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Thorndale'), 'Placement Location - City'] = 'London'

    # China
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - State'] == 'Hubei Province') & (citizenship_after_drop['Placement Location - City'] == 'Xiaogan'), 'Placement Location - City'] = 'Xiaoganzhan'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Lingang'), 'Placement Location - City'] = 'Shanghai'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - State'] == 'Hei Longjiang') & (citizenship_after_drop['Placement Location - City'] == 'Da Qing'), 'Placement Location - City'] = 'Daqing'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Beikking'), 'Placement Location - City'] = 'Beijing'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Chongzhou'), 'Placement Location - City'] = 'Chengdu'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Sichuan'), 'Placement Location - City'] = 'Chengdu'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'JINAN'), 'Placement Location - City'] = 'Jinan'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Yunnan, Guiyang'), 'Placement Location - City'] = 'Guiyang'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - Country'] == 'Hong Kong'), 'Placement Location - City'] = 'Hong Kong'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Hang Zhou'), 'Placement Location - City'] = 'Hangzhou'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Nanching'), 'Placement Location - City'] = 'Nanchang'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Guangdong'), 'Placement Location - City'] = 'Guangzhou'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Guang Zhou'), 'Placement Location - City'] = 'Guangzhou'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'GuangZhou'), 'Placement Location - City'] = 'Guangzhou'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Jiangxi'), 'Placement Location - City'] = 'Nanchang'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - City'] == 'Xuanwu, Nanjing'), 'Placement Location - City'] = 'Nanjing'
    citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - Country'] == 'Taiwan'), 'Placement Location - City'] = 'Taipei'


    problem_city = []
    problem_city_row = []
    for col_name, col_data in citizenship_after_drop.iterrows():

        if col_data['Placement Location - City'] not in dict_worldcities[col_data['Placement Location - Country']]:
            if col_data['Placement Location - State'] == "Queensland":
                citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - State'] == 'Queensland') & (citizenship_after_drop['Placement Location - City'] == col_data['Placement Location - City']), 'Placement Location - City'] = 'Brisbane'
            elif col_data['Placement Location - State'] == "New South Wales":
                citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - State'] == 'New South Wales') & (citizenship_after_drop['Placement Location - City'] == col_data['Placement Location - City']), 'Placement Location - City'] = 'Sydney'
            elif col_data['Placement Location - State'] == "NSW":
                citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - State'] == 'NSW') & (citizenship_after_drop['Placement Location - City'] == col_data['Placement Location - City']), 'Placement Location - City'] = 'Sydney'
            elif col_data['Placement Location - State'] == "Victoria":
                citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - State'] == 'Victoria') & (citizenship_after_drop['Placement Location - City'] == col_data['Placement Location - City']), 'Placement Location - City'] = 'Melbourne'
            elif col_data['Placement Location - State'] == "VIC":
                citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - State'] == 'VIC') & (citizenship_after_drop['Placement Location - City'] == col_data['Placement Location - City']), 'Placement Location - City'] = 'Melbourne'
            elif col_data['Placement Location - State'] == "Melbourne, VIC":
                citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - State'] == 'Melbourne, VIC') & (citizenship_after_drop['Placement Location - City'] == col_data['Placement Location - City']), 'Placement Location - City'] = 'Melbourne'
            elif col_data['Placement Location - State'] == "Tasmania":
                citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - State'] == 'Tasmania') & (citizenship_after_drop['Placement Location - City'] == col_data['Placement Location - City']), 'Placement Location - City'] = 'Hobart'
            elif col_data['Placement Location - State'] == "Northern Territory":
                citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - State'] == 'Northern Territory') & (citizenship_after_drop['Placement Location - City'] == col_data['Placement Location - City']), 'Placement Location - City'] = 'Darwin'
            elif col_data['Placement Location - State'] == "Western Australia":
                citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - State'] == 'Western Australia') & (citizenship_after_drop['Placement Location - City'] == col_data['Placement Location - City']), 'Placement Location - City'] = 'Perth'
            elif col_data['Placement Location - State'] == "South Australia":
                citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - State'] == 'South Australia') & (citizenship_after_drop['Placement Location - City'] == col_data['Placement Location - City']), 'Placement Location - City'] = 'Adelaide'
            elif col_data['Placement Location - State'] == "ACT":
                citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - State'] == 'ACT') & (citizenship_after_drop['Placement Location - City'] == col_data['Placement Location - City']), 'Placement Location - City'] = 'Canberra'
            elif col_data['Placement Location - State'] == "Australian Capital Territory":
                citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - State'] == 'Australian Capital Territory') & (citizenship_after_drop['Placement Location - City'] == col_data['Placement Location - City']), 'Placement Location - City'] = 'Canberra'
            elif col_data['Placement Location - Country'] == "New Zealand":
                citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - Country'] == 'New Zealand') & (citizenship_after_drop['Placement Location - City'] == col_data['Placement Location - City']), 'Placement Location - City'] = 'Auckland'
            elif col_data['Placement Location - Country'] == "Indonesia":
                citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - Country'] == 'Indonesia') & (citizenship_after_drop['Placement Location - City'] == col_data['Placement Location - City']), 'Placement Location - City'] = 'Jakarta'
            elif col_data['Placement Location - Country'] == "China":
                citizenship_after_drop.loc[(citizenship_after_drop['Placement Location - Country'] == 'China') & (citizenship_after_drop['Placement Location - City'] == col_data['Placement Location - City']), 'Placement Location - City'] = 'Beijing'  
            
                
            problem_city.append(col_data['Placement Location - City'])
            problem_city_row.append(col_data)



    # Using matching algo to fix human error in city
    for col_name, col_data in citizenship_after_drop.iterrows():
        
        ID = col_data["ID#"]
        
        tmp_city = col_data['Placement Location - City']

        if tmp_city not in dict_worldcities[col_data['Placement Location - Country']]:
            
            if col_data['Placement Location - State'] in dict_worldcities[col_data['Placement Location - Country']]:
                citizenship_after_drop.loc[(citizenship_after_drop['ID#'] == ID), 'Placement Location - City'] = col_data['Placement Location - State']
                add_entry(log_book, ID, tmp_city, col_data['Placement Location - Country'], col_data['Placement Location - State'], col_data['Placement Location - Country'])
                
            else: 
                if col_data['Placement Location - Country'] in high_freq_dic:
                    best_match = match_city_to_country(tmp_city, col_data['Placement Location - Country'], high_freq_dic)
                else:
                    best_match = match_city_to_country(tmp_city, col_data['Placement Location - Country'], dict_worldcities)
                
                if best_match is not None:
                    citizenship_after_drop.loc[(citizenship_after_drop['ID#'] == ID), 'Placement Location - City'] = best_match[1]
                    add_entry(log_book, ID, tmp_city, col_data['Placement Location - Country'], best_match[1], col_data['Placement Location - Country'])
                else:
                    add_entry(log_book, ID, tmp_city, col_data['Placement Location - Country'], 'No matching city', col_data['Placement Location - Country'])



    #from collections import Counter
    #my_counter = Counter(problem_city)

    #sorted_list = sorted(my_counter, key=lambda x: my_counter[x], reverse=True)
    #print(sorted_list)
    #print(len(sorted_list))


    citizenship_after_drop = citizenship_after_drop.drop(citizenship_after_drop.loc[citizenship_after_drop["Placement Location - City"] == 'Maputo'].index)
    citizenship_after_drop = citizenship_after_drop.dropna(subset=['Placement Location - City'])


    # Change the type of datetime
    citizenship_after_drop['Placement End Date'] = pd.to_datetime(citizenship_after_drop['Placement End Date'])
    citizenship_after_drop['Placement Start Date'] = pd.to_datetime(citizenship_after_drop['Placement Start Date'])
    citizenship_after_drop['Year'] = pd.to_datetime(citizenship_after_drop['Year'], format = '%Y')

    # Add a column called duration
    citizenship_after_drop['duration'] = citizenship_after_drop['Placement End Date'] - citizenship_after_drop['Placement Start Date'] 


    # Column of citizenship
    problem_citizenship = []

    for col_name, col_data in citizenship_after_drop.iterrows():
        if col_data['Country of Citizenship'] not in dict_worldcities:
            problem_citizenship.append(col_data['Country of Citizenship'])



    from collections import Counter
    my_counter = Counter(problem_citizenship)

    sorted_list = sorted(my_counter, key=lambda x: my_counter[x], reverse=True)


    citizenship_after_drop.loc[(citizenship_after_drop['Country of Citizenship'] == 'Viet Nam'), 'Country of Citizenship'] = 'Vietnam'
    citizenship_after_drop.loc[(citizenship_after_drop['Country of Citizenship'] == 'Republic of Korea'), 'Country of Citizenship'] = 'South Korea'
    citizenship_after_drop.loc[(citizenship_after_drop['Country of Citizenship'] == 'United States of America'), 'Country of Citizenship'] = 'United States'


    # Add a column to determine replace back to home country or not
    # Yes: back to home country
    # No: distination is not home country
    citizenship_after_drop['determine replace back to home country'] = citizenship_after_drop.apply(
        lambda row: 'yes' if row['Placement Location - Country'] == row['Country of Citizenship'] else 'no',
        axis=1
    )

    citizenship_after_drop['duration'] = citizenship_after_drop['duration'].dt.days.astype(int)
    bins = [-1, 14, 30, 90, 180, citizenship_after_drop['duration'].max()]
    labels = ['2weeks', '1month', '3months', 'half_year', 'one year or more']
    citizenship_after_drop['Duration_Category'] = pd.cut(citizenship_after_drop['duration'], bins=bins, labels=labels)

    # Produce the dataset after processing
    final_table = citizenship_after_drop[['ID#','Year', 'Program', 'Placement Location - City', 'Placement Location - State', 'Placement Location - Country', 'Placement Start Date', 'Placement End Date', "Country of Citizenship", 'duration','determine replace back to home country', 'Duration_Category']]

    # 保存处理后的Excel文件
    output_file_1 = filedialog.asksaveasfilename(defaultextension=".xlsx", title="Save excel File", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    output_file_2 = filedialog.asksaveasfilename(defaultextension=".txt", title="Save Log File", filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
    if output_file_1:
        final_table.to_excel(output_file_1, index=False)
        
    if output_file_2:
        with open(output_file_2, "w") as log_file:
            log_file.writelines(log_content)
        lbl_status.config(text="cleaned excel and logbook is saved")
        app.after(2000, app.quit)  # 2秒后关闭窗口

def load_file_1():
    global file_path_1
    file_path_1 = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if file_path_1 and file_path_2:
        lbl_status.config(text="Processing...")
        app.after(100, process_excel, file_path_1, file_path_2) 

def load_file_2():
    global file_path_2
    file_path_2 = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if file_path_1 and file_path_2:
        lbl_status.config(text="Processing...")
        app.after(100, process_excel, file_path_1, file_path_2) 

def save_log_file():
    log_file_path = filedialog.asksaveasfilename(defaultextension=".txt", title="Save Log File", filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
    if log_file_path:
        with open(log_file_path, "w") as log_file:
            log_file.writelines(log_content)  # 这里直接使用 writelines

app = tk.Tk()
app.title("Files Processor")
app.geometry("350x150")

file_path_1 = ""
file_path_2 = ""

load_button_1 = tk.Button(app, text="Load faculty Excel File", command=load_file_1)
load_button_1.pack(pady=10)

load_button_2 = tk.Button(app, text="Load worldcities Excel File", command=load_file_2)
load_button_2.pack(pady=10)

lbl_status = tk.Label(app, text="")
lbl_status.pack(pady=10)

app.mainloop()
