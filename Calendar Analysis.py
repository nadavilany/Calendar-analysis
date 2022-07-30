import pandas as pd
from datetime import timedelta
from matplotlib import pyplot as plt
import tkinter
from tkinter import *
from tkinter import ttk


def sort_topics_flex(subject, df_keys):  # sorting by subject from an editable excel file of key words
    try:
        for col in df_keys.iteritems():  # iterating through the key words and trying to sort by subject
            for i in range(len(col[1])):
                if str(col[1][i]) in str(subject):
                    return str(col[0])
        return "Manual sorting needed"  # couldnt find a key word, manual sorting needed
    except TypeError:
        return "Delete"


def category_sort(category,df_keys):  # inside and outside sorting
    try:
        for col in df_keys.iteritems():  # iterating through the key words and trying to sort by category
            for i in range(len(col[1])):
                if str(col[1][i]) in str(category):
                    return(str(col[0]))
        return "Manual sorting needed"
    except TypeError:
        return "Delete"


def build_dict(df, d, subject, denom):  # builds the output dictionary
    d[subject] = df[(df["Sorted subject"] == subject)]["Duration"].sum() / denom * 100


def timecheck(start, end):  # checks the duration of every meeting
    try:
        start = start.split(":")
        end = end.split(":")
        t1 = timedelta(hours=int(start[0]), minutes=int(start[1]))
        t2 = timedelta(hours=int(end[0]), minutes=int(end[1]))
        time = t2 - t1
        time = str(time)
        time = time.split(":")
        duration = (int(time[0]) * 60) + int(time[1])
        return duration
    except ValueError:
        return 24



def extract_month(date):  # extract month from date
    return int(str(date).split('/')[1])


def get_start_month(df):  # Helping function that helps separating the data to months, handles the year changes
    d = {}
    lst = []
    mini = df["Month"].min()
    for month in df["Month"]:
        if not month in lst:
            d[month] = d.get(month, 0) + 1
    for key, val in d.items():
        if d[key] < 5:
            lst.append(key)
    for elem in lst:
        d.pop(elem)
    if sorted(list(d.keys())) == [1, 11, 12]:
        return 11, 12, 1
    elif sorted(list(d.keys())) == [1, 2, 12]:
        return 12, 1, 2
    else:
        return mini, mini + 1, mini + 2


def analysis(df, start_month, df_keys):  # analysis over a month
    num_of_relevant = df[(df["Sorted subject"] != "Delete")]["Subject"].count()
    left_out = df[(df["Sorted subject"] == "Manual sorting needed") & (df["Category sort"] != "Outside meeting")]["Subject"].count()
    out = df[(df["Category sort"] == "Outside meeting")]["Duration"].sum()  # dealing with outside meetings
    df = df[df["Category sort"] != "Outside meeting"]
    subjected = df[(df["Sorted subject"] != "Manual sorting needed")]["Duration"].sum()
    total_time = subjected + out  # getting the denominator for percentage results
    subject_dict = {"Outside meeting": out / total_time * 100}  # building a dataframe and excel sheet of the results by percentage
    lst = list(df_keys.columns)
    for elem in lst:
        build_dict(df, subject_dict, str(elem), total_time)
    subject_dict["Percentage of success"] = (100 - (left_out / num_of_relevant * 100))
    df_subjects = pd.DataFrame(data=subject_dict, index=[start_month])
    return df_subjects


def plot(df, month_lst):  # plotting the results
    for i in range(2):
        month_lst[i] = str(month_lst[i])
    x = month_lst
    finds_mean = df["Percentage of success"].mean()
    df = df.drop("Percentage of success", 1)
    for col in df.columns:
        lst = [df[col][int(month_lst[0])], df[col][int(month_lst[1])], df[col][int(month_lst[2])]]
        if col == "Delete":
            continue
        plt.plot(x, lst, label=col)
    plt.legend(loc=2, ncol=3)
    plt.tight_layout()
    plt.title("Calendar Analysis, Percentage of success is: "+str(round(finds_mean, 1)))
    plt.xlabel("Month")
    plt.ylabel("Time spent (%/minutes)")
    plt.subplots_adjust(left=0.09, bottom=0.09)
    plt.show()
    plt.savefig("Plot", bbox_inches="tight")


def main():
    # reading the calendar
    df = pd.read_csv("input.CSV")
    df_keys = pd.read_excel("keywords.xlsx", index_col=None)
    writer = pd.ExcelWriter("output.xlsx")
    try:
        # from now on its cleaning
        lst = ['Show time as', 'Sensitivity', 'Private', 'Priority', 'Mileage',
               'Location'
            , 'Description', 'Billing Information', 'Meeting Resources', 'Optional Attendees', 'Required Attendees',
               'Reminder Time', 'Reminder Date', 'Reminder on/off','Meeting Organizer']
        df = df.drop(lst, 1)  # cleaning the db
        df = df[df['All day event'] == False]
        df = df.drop('All day event', 1)
        # From now on its sorting
        df["Duration"] = df.apply(lambda row: timecheck(row['Start Time'], row['End Time']), 1)  # creating duration series
        df["Sorted subject"] = df.apply(lambda row: sort_topics_flex(row['Subject'], df_keys), 1)  # sorting by subject
        df["Category sort"] = df.apply(lambda row: category_sort(row["Categories"],df_keys),
                                       1)  # sorting by categories
        df = df[(df["Sorted subject"] != "Delete")]  # erasing the irrelevant meetings
        df = df[(df["Category sort"] != "Delete")]
        # Splitting to months
        df["Month"] = df.apply(lambda row: extract_month(row["Start Date"]), 1)
        start_month = df["Month"].min()  # Intializing the starting month to the smallest month in the db
        start_month, middle_month, end_month = get_start_month(df)
        df1 = df[(df["Month"] == start_month)]
        df2 = df[(df["Month"] == middle_month)]
        df3 = df[(df["Month"] == end_month)]
        # Analysis for every month
        first_month_results = analysis(df1, start_month, df_keys)
        second_month_results = analysis(df2, middle_month, df_keys)
        third_month_results = analysis(df3, end_month, df_keys)
        df_subjects = pd.concat([first_month_results, second_month_results, third_month_results])
        df_subjects=df_subjects.drop('Delete',1)
        df_subjects.to_excel(writer, sheet_name="Results", index=True)
        df.to_excel(writer, sheet_name="Draft for manual sorting", index=True)
        plot(df_subjects, [start_month, middle_month, end_month])
        print("Analysis done")
    except PermissionError:
        print("Close the output file")
    finally:
        writer.close()


try:
    root = Tk()
    root.title("Calendar Analysis")

    mainframe = ttk.Frame(root, padding="20 20 20 20")
    mainframe.grid(column=1, row=1, sticky=(N, W, E, S))
    mainframe.columnconfigure(0, weight=1)
    mainframe.rowconfigure(0, weight=1)

    label = tkinter.Label(mainframe, text="Welcome to Calendar Analysis").grid()
    button = ttk.Button(mainframe, text='Press to start', command=lambda: main())
    button.grid()
    root.mainloop()
except PermissionError:
    print("Close the output file")
