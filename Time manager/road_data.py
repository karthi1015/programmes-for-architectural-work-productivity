import pandas as pd
import sqlite3
import matplotlib.pyplot as plt
import operator
import numpy as np
from datetime import datetime, timedelta, date
# import csv_to_db

conn = sqlite3.connect('schedule.db')
c = conn.cursor()
# csv_to_db.csv_to_db()
df = pd.read_sql('SELECT * FROM engenium_projects', conn)

def hours_spent_func(select_by):
    seconds = 0
    for ea_second in df[select_by]['hours_spent']:
        seconds = seconds + ea_second
    mins, seconds = divmod(seconds, 60)
    hours, mins = divmod(mins, 60)
    return('%02d:%02d:%02d' % (hours, mins, seconds))

def hours_spent_func_unit(select_by):
    seconds = 0
    for ea_second in df[select_by]['hours_spent']:
        seconds = seconds + ea_second
    return(seconds/3600)

def print_time_code(select_by):
    print('Time code used')
    return(df[select_by]['time_code'].drop_duplicates().to_string(index=False))

def start_date_format(year_input, month_input, day_input):
    tdelta = timedelta(days=6)
    st_date = date(year_input, month_input, day_input)
    end_date = st_date + tdelta
    return(str(st_date),str(end_date))

def check_hours():
    start_date_input = str(input("Week starting date (e.g 2017-7-1) : "))

    full_date_input = datetime.strptime(start_date_input, '%Y-%m-%d')
    year = int(full_date_input.year)
    month = int(full_date_input.month)
    day = int(full_date_input.day)
    start_date_from_db = df[df['start_date'] == start_date_format(year_input=year, month_input=month, day_input=day)[0]]
    print(start_date_format(year,month,day)[0], start_date_format(year,month,day)[1])

    start_to_end = (df['start_date'] >= start_date_format(year,month,day)[0]) & (df['start_date'] <= start_date_format(year,month,day)[1])
    selected_df = df.loc[start_to_end]

    print('Total hours worked: ',hours_spent_func(start_to_end))

def autolabel(rects):
    lst = []
    for rect in rects:
        height = rect.get_height()
        ax.text(rect.get_x() + rect.get_width()/2. , height, '%s' % height, ha = 'center', va = 'bottom')

base_input = str(input("1.Project base / 2.Hours for a week: "))

if base_input == '1':
    project_number_input = str(input("Project number: ")).upper()
    print('')

    p_num = df[df['project_number'] == project_number_input]
    category_time_activity = p_num[['categories','time_code','activity']]
    print('PROJECT NAME')
    print(p_num['project_name'].drop_duplicates().to_string(index=False))
    print('')
    print(category_time_activity.drop_duplicates().to_string(index=False))
    print('')

    p_num_cat = df[(df['project_number'] == project_number_input) & (df['categories'] == 'ARCHITECTURAL')]
    first_date = p_num_cat['start_date'].min()
    up_to_date = p_num_cat['start_date'].max()
    print('Date from ', first_date, 'to ' ,up_to_date, ' for ARCHITECTURAL CATEGORY ONLY.')
    print('')

    how_to_show = str(input("How do you want to filter project? 1.Category/Time code/Activities 2.Categories 3.Time Code 4.Activity graph "))

    if how_to_show == '1':
        total_input = str(input("Category(A,C,S,6,7,8)/TimeCode/Activity: "))
        print('')
        categories_input = total_input.split('/')[0].upper()
        time_code_input = total_input.split('/')[1].upper()
        activity_input = total_input.split('/')[2]
        if categories_input == 'A':
            categories_input = 'ARCHITECTURAL'
        elif categories_input == 'C':
            categories_input = 'CONCEPT'
        elif categories_input == 'S':
            categories_input = 'STRUCTURAL'
        elif categories_input == '6':
            categories_input == '6T'
        elif categories_input == '7':
            categories_input == '7T'
        elif categories_input == '8':
            categories_input == '8CODE'
        else:
            pass
        select_by = (df['project_number'] == project_number_input) & (df['categories'] == categories_input) & (df['time_code'] == time_code_input) & (df['activity'] == activity_input)
        print('Time spent')
        print(hours_spent_func(select_by))

    elif how_to_show == '2':
        categories_input = str(input("Category(A,C,S,6,7,8): ")).upper()
        print('')
        if categories_input == 'A':
            categories_input = 'ARCHITECTURAL'
        elif categories_input == 'C':
            categories_input = 'CONCEPT'
        elif categories_input == 'S':
            categories_input = 'STRUCTURAL'
        elif categories_input == '6':
            categories_input == '6T'
        elif categories_input == '7':
            categories_input == '7T'
        elif categories_input == '8':
            categories_input == '8CODE'
        else:
            pass
        select_by = (df['project_number'] == project_number_input) & (df['categories'] == categories_input)
        print(print_time_code(select_by))
        print('Time spent')
        print(hours_spent_func(select_by))

    elif how_to_show == '3':
        time_code_input = str(input("Time_code: ")).upper()
        print('')
        select_by = (df['project_number'] == project_number_input) & (df['time_code'] == time_code_input)
        print('Time spent')
        print(hours_spent_func(select_by))

    elif how_to_show == '4':
        dict_graph = {}
        for index, row_row in p_num[['activity']].drop_duplicates().iterrows():
            select_by = (df['project_number'] == project_number_input) & (df['activity'] == row_row['activity'])
            dict_graph[row_row['activity']] = hours_spent_func_unit(select_by)

        x = [a for a in range(len(dict_graph))]
        y = [v for k,v in dict_graph.items()]

        fig, ax = plt.subplots()
        rects_1 = ax.bar(x, y, 0.8)

        # np.arange(len(x))  =  x

        my_xticks = [k for k,v in dict_graph.items()]
        plt.xticks(x, my_xticks, rotation='vertical')
        plt.xlabel('Activities')
        plt.ylabel('Hours')
        plt.title('Time Spent')

        autolabel(rects_1)
        plt.show()



# ''' SHOW Y VALUE AS HOURS
#         # sorted_dict_graph = sorted(dict_graph.items(), key=operator.itemgetter(1))
#
#         # x = [ a for a in range(len(sorted_dict_graph))]
#         # y = [ b for b in range(len(sorted_dict_graph))]
#         # x_lst = []
#         # y_lst = []
#
#         # my_xticks = [sorted_dict_graph[i][0] for i in range(len(sorted_dict_graph))]
#         # my_yticks = [sorted_dict_graph[j][1] for j in range(len(sorted_dict_graph))]
#         # plt.yticks(y, my_yticks)
# '''

elif base_input == '2':
    check_hours()


c.close()
conn.close()
