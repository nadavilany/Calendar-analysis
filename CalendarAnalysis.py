import numpy as np
import pandas as pd
from datetime import timedelta
from matplotlib import pyplot as plt
import tkinter
from tkinter import *
from tkinter import ttk


def sort_topics_flex(subject, df_keys):  # Sorting by subject from an editable excel file of key words
    try:
        for col in df_keys.iteritems():  # Iterating through the key words and trying to sort by subject
            for i in range(len(col[1])):
                if str(col[1][i]) in str(subject):
                    return str(col[0])
        return "Manual sorting needed"  # Couldnt find a key word, manual sorting needed
    except TypeError:
        return "Delete"


def category_sort(category, df_keys):  # Sorting by category
    try:
        if str(category) == "nan":  # Deals with uncategorized meetings
            return "Manual sorting needed"
        for col in df_keys.iteritems():  # Iterating through the key words and trying to sort by category
            for i in range(len(col[1])):
                if str(col[1][i]) in str(category):
                    return (str(col[0]))
        return "Manual sorting needed"
    except TypeError:
        return "Delete"


def build_dict(df, d, subject, denom):  # Building the output dictionary
    d[subject] = df[(df["Sorted category"] == subject)][
                     "Duration"].sum() / denom * 100  # Prioritizing sorting by category
    # Adding sorting by subject
    d[subject] += df[(df["Sorted subject"] == subject) & (df["Sorted category"] == "Manual sorting needed")][
                      "Duration"].sum() / denom * 100


def timecheck(start, end):  # Checking the duration of every meeting
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


def extract_month(date):  # Extracting month from date
    return int(str(date).split('/')[1])


def get_start_month(df):  # Separating the data to months, handles the year changes
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


def analysis(df, start_month, df_keys):  # Analysis over a month
    num_of_relevant = df[(df["Sorted subject"] != "Delete") & (df["Sorted category"] != "Delete")][
        "Subject"].count()  # Counting all the relevant meetings
    left_out = \
        df[(df["Sorted subject"] == "Manual sorting needed") & (df["Sorted category"] == "Manual sorting needed")][
            "Subject"].count()  # Counting the number of meetings we couldnt sorted
    subjected = \
        df[(df["Sorted subject"] != "Manual sorting needed") | (df["Sorted category"] != "Manual sorting needed")][
            "Duration"].sum()
    subject_dict = {}  # Building a dataframe and excel sheet of the results by percentage
    lst = list(df_keys.columns)
    for elem in lst:
        build_dict(df, subject_dict, str(elem), subjected)
    subject_dict["Success rate"] = (100 - (left_out / num_of_relevant * 100))
    df_subjects = pd.DataFrame(data=subject_dict, index=[start_month])
    return df_subjects


def plot(df, month_lst):  # Plotting the results
    for i in range(2):  # Prepering months
        month_lst[i] = str(month_lst[i])
    x = month_lst
    finds_mean = df["Success rate"].mean()  # Finding the percentage of success
    df = df.drop("Success rate", axis=1)
    # Linear chart
    fig, ax = plt.subplots()
    for col in df.columns:  # Plotting for each subject
        lst = [df[col][int(month_lst[0])], df[col][int(month_lst[1])], df[col][int(month_lst[2])]]
        if col == "Delete":
            continue
        ax.plot(x, lst, label=col)
    fig.legend(loc=4, ncol=3)
    plt.tight_layout()
    plt.title("Calendar Analysis")
    plt.xlabel("Month")
    plt.ylabel("Time spent (%/minutes)")
    fig.subplots_adjust(left=0.09, bottom=0.25)
    props = dict(boxstyle='round', facecolor='wheat', alpha=0.5)
    # place a text box in upper left in axes coords
    fig.text(0.04, 0.1, "success rate: " + str(round(finds_mean, 1)), fontsize=10,
             verticalalignment='top')
    plt.style.use('ggplot')
    plt.savefig("Linear Plot", bbox_inches="tight")
    # Bar chart
    fig1, ax1 = plt.subplots()
    for i in range (len(x)): # Prepering the x axis
        x[i]=int(x[i])
    plt.xticks(x)
    t=np.array(x)
    lst3=[]
    for col in df.columns:
        lst2=[]
        for month in month_lst:
            lst2+=([df[col][int(month)]])
        lst3.append(lst2)
    color_lst=['r','b','g','c','k','y','m','0.5']
    for i in range (len(lst3)):
        ax1.bar(t - (len(lst3)//2)/10+i/10, lst3[i], color=color_lst[i%(len(lst3))], width=0.1,label=df.columns[i])

    for p in ax1.patches:
        ax1.text(p.get_x() + p.get_width()/2, p.get_height(), int(p.get_height()),
                fontsize=10, color='black', ha='center', va='bottom')
    plt.style.use('ggplot')
    fig1.legend(loc=4, ncol=3, fancybox=True)
    plt.tight_layout()
    plt.title("Calendar Analysis")
    plt.xlabel("Month")
    plt.ylabel("Time spent (%/minutes)")
    fig1.subplots_adjust(left=0.09, bottom=0.25)
    props = dict(boxstyle='round', facecolor='wheat', alpha=0.5)
    # place a text box in upper left in axes coords
    fig1.text(0.04, 0.1, "success rate: " + str(round(finds_mean, 1)), fontsize=10,
              verticalalignment='top')
    plt.savefig("Bar chart", bbox_inches="tight")


def clean_db(df):  # Cleaning the db by dropping irrelevant columns
    lst = ['Show time as', 'Sensitivity', 'Private', 'Priority', 'Mileage',
           'Location'
        , 'Description', 'Billing Information', 'Meeting Resources', 'Optional Attendees', 'Required Attendees',
           'Reminder Time', 'Reminder Date', 'Reminder on/off', 'Meeting Organizer']
    df = df.drop(lst,axis=1)
    df = df[df['All day event'] == False]
    df = df.drop('All day event', axis=1)


def sort_db(df, df_keys):  # Sorting each meeting by category and subject
    df["Duration"] = df.apply(lambda row: timecheck(row['Start Time'], row['End Time']), 1)  # creating duration series
    df["Sorted subject"] = df.apply(lambda row: sort_topics_flex(row['Subject'], df_keys), 1)  # sorting by subject
    df["Sorted category"] = df.apply(lambda row: category_sort(row["Categories"], df_keys), 1)  # sorting by categories
    df = df[(df["Sorted subject"] != "Delete")]  # erasing the irrelevant meetings
    df = df[(df["Sorted category"] != "Delete")]
    return df


def analysis_by_month(df, df_keys):  # Splitting by months, running the analysis process and ploting
    # Splitting to months
    df["Month"] = df.apply(lambda row: extract_month(row["Start Date"]), 1)
    start_month, middle_month, end_month = get_start_month(df)
    df1 = df[(df["Month"] == start_month)]
    df2 = df[(df["Month"] == middle_month)]
    df3 = df[(df["Month"] == end_month)]
    # Analysis for every month
    first_month_results = analysis(df1, start_month, df_keys)
    second_month_results = analysis(df2, middle_month, df_keys)
    third_month_results = analysis(df3, end_month, df_keys)
    df_results = pd.concat([first_month_results, second_month_results, third_month_results])
    df_results = df_results.drop('Delete', axis=1)
    plot(df_results, [start_month, middle_month, end_month])
    return df_results


def main():
    try:
        # Reading the calendar
        df = pd.read_csv("input.CSV")
        df_keys = pd.read_excel("keywords.xlsx", index_col=None)
        writer = pd.ExcelWriter("output.xlsx")
        clean_db(df)  # cleaning the db
        df = sort_db(df, df_keys)  # Sorting each meeting by category and subject
        df_results = analysis_by_month(df, df_keys)  # Splitting by months, running the analysis process and ploting
        # Writing to excel
        df_results.to_excel(writer, sheet_name="Results", index=True)
        df.to_excel(writer, sheet_name="Draft for manual sorting", index=True)
        print("Analysis done")
        writer.close()
    except PermissionError:
        print("Please close the output file")


try:  # Button and GUI
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
    print("Please close the output file")
