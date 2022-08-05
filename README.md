# Calendar-analysis
Semi automatic calendar analysis tool

This python tool works on a raw csv file of outlook calendar (3 months span) to analyze each meeting on the calendar by a user managed editable excel file.

Workflow for users: (see tutorial branch)
  1. Update the excel file named "keywords" to ensure high quality results
  2. Export A csv file from outlook in a span of 3-month, name it as "input" and save it in the same location of the script.
  3. Run the script (option for exe file creation), there will be a pop-up button to start the analysis.
  4. Press to start.
  5. There will be a pop-up plot that shows the results of the analysis over the last 3 months and saves itself in the same location of the script as "plot.PNG".
  6. New excel file is created name "output.xlsx" with 2 sheets:
    6.1. "Results" - shows the numeric results of the analysis that are shown in the plot.
    6.2. "Draft for manual sorting" - shows the data frame after the analysis, for the user to check the analysis process and continue manually.
  7. In the Headline of the plot created and in the excel file of "output" there will be a documentation of the success rate of the analysis. if this percentage is          below 70%, we recommend to update the excel file of "keywords" in correlation to the calendar subjects and categories in outlook.

Added files:
  1. PyCharm file - raw code.
  2. Application file - EXE.
  3. tutorial file - ppt file that explains how to use the tool and contains Q&A segment.
  4. Input file for example.
  5. Keywords file for example.
  6. Output and plots for example.

Algorithm's workflow:
  1. Reads the Input and Keywords files to a pandas dataframe.
  2. Cleans the input db, drops irrelevant columns.
  3. For each meeting:
    3.1 Creates "Duration" series from "Start Time" colum and "End Time" colum.
    3.2 Creates "Sorted Subject" series from "Subject" colum by matching a keywords from "Keywords" file.
    3.3 Creates "Sorted Category" series from "Categories" colum by matching a keywords from "Keywords" file.
  4. Analyzes the relative percantege for each colum of "keywords" file (first by "Sorted Category" and than by "Sorted Subject") and calculates the success rate of        the analysis process (failure of analysis happens when there was no success in sorting by subject nor by category)
  5. Plots as linear by months and as bars by subjects and category sorted.

Clarifications:
  1. Calendar Analysis tool works on an export from desktop version of Microsoft Outlook's calendar.
  2. For period of over 3 months, the tool will work on the first 3 months of this period.
  3. This tool is designed to enable a non-coding user to analyze a calendar by editing the file "Keywords" only.
  4. In case of a clash between "Sorted Subject" and "Sorted Category", the algorithm will prioritize "Sorted Category".

Future upgrades and appreciated contributions:
  1. Full automatization upgrades:
    1.1. Developing an abillity to export the calendar directly from Microsoft Outlook API
    1.2. Creating a matching GUI and UI for filtering months.
  2. Flexability upgrades:
    2.1. Enabling analysis for more or less than 3 months.
    2.2. Enabling analysis between non-following months.
  3. Complexity and efficiency upgrades.
  4. UI and aesthetics upgrades.
