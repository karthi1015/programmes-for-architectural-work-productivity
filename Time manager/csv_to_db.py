from datetime import datetime
from csv import DictReader
import os
import sqlite3

### ON MAC
# userhome = os.path.expanduser('~')
# csv_file= userhome + r'/Desktop/OneDrive/Documents/Programming/Python3/Practice/Time_manage/timesheet.csv'
# reader = DictReader(open(csv_file))

### ON WINDOW
reader = DictReader(open('timesheet.csv'))

# conn = sqlite3.connect(':memory:')
conn = sqlite3.connect('schedule.db')
c = conn.cursor()

c.execute("""CREATE TABLE IF NOT EXISTS engenium_projects(
            project_number TEXT,
            project_name TEXT,
            categories TEXT,
            time_code TEXT,
            activity TEXT,
            start_date TEXT,
            start_time TEXT,
            end_time TEXT,
            hours_spent INTEGER,
            description TEXT
            )""")
conn.commit()

sb = 'Subject'
ct = 'Categories'
s_date = 'Start Date'
s_time = 'Start Time'
e_time = 'End Time'
dsc = 'Description'

def time_spent():
    start_time_lst = row[s_time].split()
    start_time = ''.join(start_time_lst)
    end_time_lst = row[e_time].split()
    end_time = ''.join(end_time_lst)
    FMT = '%I:%M:%S%p'
    start_convert = datetime.strptime(start_time, FMT)
    end_convert = datetime.strptime(end_time, FMT)
    time_delta = end_convert - start_convert
    time_second = int(time_delta.seconds)
    return(time_second)

def all_to_data():
    project_number_to_data = row[sb].split('_')[0].upper()
    project_name_to_data = row[sb].split('_')[2].upper()
    category_to_data = row[ct].upper()
    time_code_to_data = row[sb].split('_')[1].upper()
    start_date_before_convert = datetime.strptime(row[s_date], '%m/%d/%Y')
    start_date_to_data = start_date_before_convert.strftime('%Y-%m-%d')
    start_time_to_data = row[s_time].upper()
    end_time_to_data = row[e_time].upper()
    description_to_data = row[dsc].lower()
    try:
        activity_to_data = row[sb].split('_')[3].lower()
    except:
        activity_to_data = None
    hours_spent_to_data = time_spent()
    c.execute("""SELECT * FROM engenium_projects WHERE (project_number=?
                        AND start_date=? AND start_time=?)""",
                        (project_number_to_data, start_date_to_data,
                        start_time_to_data))
    entry = c.fetchone()
    if entry is None:
        with conn:
            c.execute("""INSERT INTO engenium_projects (project_number, project_name,
                                categories, time_code, activity,
                                start_date, start_time, end_time,
                                hours_spent, description)
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                                (project_number_to_data, project_name_to_data, category_to_data,
                                time_code_to_data, activity_to_data,
                                start_date_to_data, start_time_to_data,
                                end_time_to_data, hours_spent_to_data,
                                description_to_data))
    else:
        with conn:
            c.execute("""UPDATE engenium_projects SET project_name=?,
                                categories=?, time_code=?, activity=?,
                                end_time=?,
                                hours_spent=?, description=? WHERE (project_number=?
                                                    AND start_date=? AND start_time=?)""",
                                                    (project_name_to_data, category_to_data,
                                                    time_code_to_data, activity_to_data,
                                                    end_time_to_data, hours_spent_to_data,
                                                    description_to_data, project_number_to_data, start_date_to_data,
                                                    start_time_to_data))

def eight_code_to_data():
    project_number_to_data = None
    category_to_data = row[ct].upper()
    time_code_to_data = row[sb].split('_')[0].upper()
    start_date_before_convert = datetime.strptime(row[s_date], '%m/%d/%Y')
    start_date_to_data = start_date_before_convert.strftime('%Y-%m-%d')
    start_time_to_data = row[s_time].upper()
    end_time_to_data = row[e_time].upper()
    description_to_data = row[dsc].lower()
    activity_to_data = row[sb].split('_')[1].lower()
    hours_spent_to_data = time_spent()
    c.execute("""SELECT * FROM engenium_projects WHERE (time_code=?
                        AND start_date=? AND start_time=?
                        )""",
                        (time_code_to_data, start_date_to_data,
                        start_time_to_data))
    entry = c.fetchone()
    if entry is None:
        with conn:
            c.execute("""INSERT INTO engenium_projects (project_number,
                                    categories, time_code,
                                    activity, start_date,
                                    start_time, end_time,
                                    hours_spent, description)
                                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                                    (project_number_to_data, category_to_data,
                                    time_code_to_data, activity_to_data,
                                    start_date_to_data, start_time_to_data,
                                    end_time_to_data, hours_spent_to_data,
                                    description_to_data))
    else:
        with conn:
            c.execute("""UPDATE engenium_projects SET
                                categories=?,activity=?,
                                end_time=?,
                                hours_spent=?, description=? WHERE (time_code=?
                                                    AND start_date=? AND start_time=?)
                                                    """,
                                                    (category_to_data, activity_to_data,
                                                    end_time_to_data, hours_spent_to_data,
                                                    description_to_data, time_code_to_data, start_date_to_data,
                                                    start_time_to_data))

for row in reader:
    if row[ct] == 'Personal Appointment':
        pass
    elif row[ct] == '8Code':
        eight_code_to_data()
    else:
        all_to_data()


c.close()
conn.close()
