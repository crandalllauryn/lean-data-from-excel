# Set up
import argparse
import os
import sys
import pandas as pd
from datetime import datetime
from matplotlib import pyplot as plt
from matplotlib import interactive
from pprint import pprint
import re

DEFAULT_EXCEL_INPUT_FILEPATH = os.path.join('Mission_2.xlsx')

# datetime object containing current date and time
now = datetime.now()
print("Today's date and time:", now)

plt.style.use('seaborn-whitegrid')

emoji_pattern = re.compile("["
                           u"\U0001F600-\U0001F64F"  # emoticons
                           u"\U0001F300-\U0001F5FF"  # symbols & pictographs
                           u"\U0001F680-\U0001F6FF"  # transport & map symbols
                           u"\U0001F1E0-\U0001F1FF"  # flags (iOS)
                           "]+", flags=re.UNICODE)

clean_task_names = []
# Functions


def get_user_inputs():
    parser = argparse.ArgumentParser(
        description="Lean data from excel file",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument('excelfile', default=DEFAULT_EXCEL_INPUT_FILEPATH,
                        help='Excel file from teams board you want to evaluate')

    args = parser.parse_args()

    if not os.path.isfile(args.excelfile):
        print(f'ERROR: {args.excelfile} not found')
        sys.exit(1)

    return args.excelfile


def read_excel_file(excelfile):
    df = pd.read_excel(excelfile, index_col=0, header=0)
    df2 = df.fillna(value=0)

    return df2


def pull_and_clean_task_names(df2):
    task_names = list(df2.iloc[:, 0])
    for i in range(len(task_names)):
        clean_task_names.append(emoji_pattern.sub(r'', task_names[i]))

    return clean_task_names


def pull_created_dates(df2):
    created_dates = list(df2.iloc[:, 6])

    return created_dates


def pull_completed_dates(df2):
    completed_dates = list(df2.iloc[:, 10])

    return completed_dates


def pull_bucket_names(df2):
    bucket_names = list(df2.iloc[:, 1])

    return bucket_names


def pull_label_data(df2):
    label_data = list(df2.iloc[:, 15])

    return label_data


def comp_to_create(task_names, create_dates, complete_dates):
    for i in range(len(task_names)):
        if complete_dates[i] == 0:
            complete_dates[i] = create_dates[i]

    return complete_dates

# WIP


def generate_table_dic(task_names, create_dates, complete_dates, bucket_names):
    table_dic = {}
    for i in range(len(task_names)):
        a = datetime.strptime(create_dates[i], '%m/%d/%Y')
        b = datetime.strptime(complete_dates[i], '%m/%d/%Y')
        cycle_time = b-a
        days = cycle_time.days
        table_dic[task_names[i]] = [bucket_names[i], days]

    return table_dic


def generate_avg_dic(task_names, create_dates, complete_dates):
    avg_dic = {}
    for i in range(len(task_names)):
        a = datetime.strptime(create_dates[i], '%m/%d/%Y')
        b = datetime.strptime(complete_dates[i], '%m/%d/%Y')
        cycle_time = b-a
        days = cycle_time.days
        avg_dic[i] = days

    return avg_dic


def generate_graph_dic(task_names, create_dates, complete_dates):
    scatter_plot_dic = {}
    for i in range(len(task_names)):
        a = datetime.strptime(create_dates[i], '%m/%d/%Y')
        b = datetime.strptime(complete_dates[i], '%m/%d/%Y')
        cycle_time = b-a
        days = cycle_time.days
        scatter_plot_dic[b] = days

    return scatter_plot_dic


def find_lean_average(avg_dic):
    avg = 0
    for val in avg_dic.values():
        avg += val

    avg = avg // len(avg_dic)

    return(avg)


def lean_data_graph(scatter_plot_dic, average):
    ax = plt.subplot()
    for k, v in scatter_plot_dic.items():
        ax.scatter(k, v,)

    ax.axhline(average, label='Average Cycle Time', linestyle='--')

    plt.legend(loc='upper left')
    plt.title('Lean Cycle Time')
    plt.xlabel('Dates')
    plt.ylabel('Cycle Time')
    plt.tight_layout()

# Main code


def main():
    excelfile = get_user_inputs()
    df2 = read_excel_file(excelfile)
    task_names = pull_and_clean_task_names(df2)
    created_dates = pull_created_dates(df2)
    completed_dates = pull_completed_dates(df2)
    bucket_names = pull_bucket_names(df2)
    label_names = pull_label_data(df2)
    completed_dates = comp_to_create(
        task_names, created_dates, completed_dates)
    table_dic = generate_table_dic(
        task_names, created_dates, completed_dates, bucket_names)
    avg_dic = generate_avg_dic(task_names, created_dates, completed_dates)
    scatter_plot_dic = generate_graph_dic(
        task_names, created_dates, completed_dates)
    average = find_lean_average(avg_dic)
    lean_data_graph(scatter_plot_dic, average)

    pprint(table_dic)
    print("Our cycle time average is : " + str(average) + ' days')

    plt.show()


# Run
if __name__ == '__main__':
    main()
