# imports

import pandas as pd

from datetime import datetime

from matplotlib import pyplot as plt

from pprint import pprint

import re

 

# current date and time

now = datetime.now()

 

# plot set up

plt.style.use('seaborn-whitegrid')

plt.rcParams["figure.autolayout"] = True

 

path = r'C:\Users\223050503\Downloads\Mission 2.xlsx'

 

# strip emojis

emoji_pattern = re.compile("["

                           u"\U0001F600-\U0001F64F"  # emoticons

                           u"\U0001F300-\U0001F5FF"  # symbols & pictographs

                           u"\U0001F680-\U0001F6FF"  # transport & map symbols

                           u"\U0001F1E0-\U0001F1FF"  # flags (iOS)

                           "]+", flags=re.UNICODE)

clean_task_name = []

 

# all label names

 

# Ident - Enter

# Constraints

# Ident - Delete

# Disco Entry

# Enter Waypoint List

# Pending GPM Run

 

# Pull Teams Excel Data

df = pd.read_excel(path, index_col=0, header=0)

df2 = df.fillna(value=0)

 

task_name = list(df2.iloc[:, 0])

for i in range(len(task_name)):

    clean_task_name.append(emoji_pattern.sub(r'', task_name[i]))

 

bucket_data = list(df2.iloc[:, 1])

pull_created_date = list(df2.iloc[:, 6])

pull_completed_date = list(df2.iloc[:, 10])

label_data = list(df2.iloc[:, 15])

 

# Pre-define your dictionaries and lists

 

done_avg_dic = {}

wip_avg_dic = {}

wip_dic = {}

 

x = []

y = []

 

# Loop to organize done data

 

for i in range(len(task_name)):

    if pull_completed_date[i] == 0:

        pull_completed_date[i] = pull_created_date[i]

    a = datetime.strptime(pull_created_date[i], '%m/%d/%Y')

    b = datetime.strptime(pull_completed_date[i], '%m/%d/%Y')

    cycle_time = b-a

    cycle_time = cycle_time.days

    if cycle_time > 0:

        if bucket_data[i] == 'Done':

            x.append(b)

            y.append(cycle_time)

            done_avg_dic[i] = cycle_time

 

# find done avg

done_avg = 0

 

for val in done_avg_dic.values():

    done_avg += val

 

done_avg = done_avg // len(done_avg_dic)

 

# plot done tasks

ax = plt.subplot()

 

plt.plot(x, y, 'o')

 

ax.axhline(done_avg, label='Average Cycle Time of Completed items', linestyle='--')

 

plt.legend(loc='upper right')

plt.title('Cycle Time Chart of Completed items')

plt.xlabel('Completion Date')

plt.ylabel('Cycle Time Hours')

 

plt.show()

 

print("Our done cycle time average is: " + str(done_avg) + " days")

 

# x = []

# y = []

 

# for next team

# for i in range(len(task_name)):

#     if pull_completed_date[i] == 0:

#         pull_completed_date[i] = pull_created_date[i]

#     a = datetime.strptime(pull_created_date[i], '%m/%d/%Y')

#     b = datetime.strptime(pull_completed_date[i], '%m/%d/%Y')

#     cycle_time = b-a

#     cycle_time = cycle_time.days

#     if cycle_time > 0 and bucket_data[i] == 'Done':

#         x.append(b)

#         y.append(cycle_time)

#         done_avg_dic[i] = cycle_time

 

# for all wip items

for i in range(len(task_name)):

    a = datetime.strptime(pull_created_date[i], '%m/%d/%Y')

    wip = now - a

    wip = wip.days

    if bucket_data[i] == 'In Work' or bucket_data[i] == 'Review':

        wip_avg_dic[i] = wip

 

# Find WIP Average Cycle Time

wip_avg = 0

 

for val in wip_avg_dic.values():

    wip_avg += val

 

wip_avg = wip_avg // len(wip_avg_dic)

 

print("Our wip cycle time average is: " + str(wip_avg) + " days")