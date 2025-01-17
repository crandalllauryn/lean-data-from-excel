# Imports
import pandas as pd
from datetime import datetime
from matplotlib import pyplot as plt
from pprint import pprint
import re

# Current date and time
now = datetime.now()

# Plot setup
plt.style.use('seaborn-whitegrid')
plt.rcParams["figure.autolayout"] = True

# Path to Excel file
path = r'C:\Users\223050503\Downloads\Mission 2.xlsx'

# Strip emojis
emoji_pattern = re.compile(
    "["
    u"\U0001F600-\U0001F64F"  # Emoticons
    u"\U0001F300-\U0001F5FF"  # Symbols & pictographs
    u"\U0001F680-\U0001F6FF"  # Transport & map symbols
    u"\U0001F1E0-\U0001F1FF"  # Flags (iOS)
    "]+", flags=re.UNICODE
)

# Initialize variables
clean_task_name = []

# Pull Teams Excel Data
df = pd.read_excel(path, index_col=0, header=0)
df2 = df.fillna(value=0)

task_name = list(df2.iloc[:, 0])
for name in task_name:
    clean_task_name.append(emoji_pattern.sub(r'', name))

bucket_data = list(df2.iloc[:, 1])
pull_created_date = list(df2.iloc[:, 6])
pull_completed_date = list(df2.iloc[:, 10])
label_data = list(df2.iloc[:, 15])

# Pre-define dictionaries
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
    cycle_time = (b - a).days

    if cycle_time > 0 and bucket_data[i] == 'Done':
        x.append(b)
        y.append(cycle_time)
        done_avg_dic[i] = cycle_time

# Calculate done average cycle time
done_avg = sum(done_avg_dic.values()) // len(done_avg_dic)

# Plot done tasks
ax = plt.subplot()
plt.plot(x, y, 'o')
ax.axhline(done_avg, label='Average Cycle Time of Completed Items', linestyle='--')
plt.legend(loc='upper right')
plt.title('Cycle Time Chart of Completed Items')
plt.xlabel('Completion Date')
plt.ylabel('Cycle Time (Days)')
plt.show()

print(f"Our done cycle time average is: {done_avg} days")

# Loop for WIP items
for i in range(len(task_name)):
    a = datetime.strptime(pull_created_date[i], '%m/%d/%Y')
    wip = (now - a).days

    if bucket_data[i] in ['In Work', 'Review']:
        wip_avg_dic[i] = wip

# Calculate WIP average cycle time
wip_avg = sum(wip_avg_dic.values()) // len(wip_avg_dic)
wip_avg = wip_avg // len(wip_avg_dic)

print("Our wip cycle time average is: " + str(wip_avg) + " days")
