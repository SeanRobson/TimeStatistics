# import module to read and write in .xlsx files
import pandas as pd
# import module to dates and time
import datetime as dt
from datetime import datetime

def load_data():
    """
    Open the excel file, get number of rows and columns, sorts entries by person then time
    """

    # open the excel data file
    input_file = pd.read_excel('lil processy.xlsx')

    # determine the size of the excel file (important to get number of rows)
    row, column = input_file.shape

    """
    Read the excel file, get the important information, change to correct data type if necessary
    """

    # create data storage array
    data = []
    denied = 0

    # iterate over each entry and store important information
    for entry in range(0, row):

        # create temp data storage array (empty each iteration)
        temp = []

        # get information
        name = input_file.iloc[entry, 0]
        ID = input_file.iloc[entry, 2]
        time = input_file.iloc[entry, 4]
        door = input_file.iloc[entry, 5]
        access = input_file.iloc[entry, 6]

        # add it all to one entry in temp
        temp = [name, ID, time, door, access]

        if "CR" in temp[3]:
            pass
        elif temp[4] != "GRANTED":
            denied += 1
        else:
            # append to good data
            data.append(temp)

    # sort data by person (x[0]) and time (x[2])
    data.sort(key=lambda x: (x[0], x[2]))

    return data, row, denied


def check_data_loaded(data):
    """
    Check to see if data has been loaded
    """
    data_check = 1

    if not data:
        data_check = 1
    else:
        data_check = 0

    return data_check


def process_data(data, row):
    """
    Determine the earliest and latest dates in the given data
    """

    # set earliest date as now
    earliest_date = dt.datetime.now()
    # set latest date as Jan 1 2000
    latest_date = dt.datetime(2000, 1, 1)

    # go through all the entries and find the earliest and latest dates
    for entry in data:

        # current entry's date
        entry_date = entry[2]

        # if the current entry's date is earlier than the earliest date
        if entry_date < earliest_date:

            # set earliest date
            earliest_date = entry_date

        # if the current entry's date is after than the latest date
        if entry_date > latest_date:

            # set latest date
            latest_date = entry_date

    # Determine the number of people who came on site
    TGen_total = 0
    CON_total = 0
    VIS_total = 0
    Other_total = 0
    index = 0

    for entry in data:
        person = entry[0]

        try:
            if person != data[index + 1][0] and "CON" in str(entry[1]):
                CON_total += 1
            elif person != data[index + 1][0] and "APP" in str(entry[1]):
                Other_total += 1
            elif person != data[index + 1][0] and entry[1] * 1 > 1:
                TGen_total += 1
            elif person != data[index + 1][0] and "V" in person:
                VIS_total += 1

            index += 1

        except IndexError:
            break

        total = TGen_total + VIS_total + Other_total + CON_total

    index = 0
    denied = 0

    tgen_total = dt.timedelta(0, 0, 0)
    con_total = dt.timedelta(0, 0, 0)
    total = dt.timedelta(0, 0, 0)

    for entry in data:
        try:
            if entry[4] != "GRANTED":
                denied += 1
                # capture if the same person
            elif "IN" in entry[3] and entry[4] == "GRANTED":
                start = entry[2]
                ID = entry[1]
                if ID == data[index + 1][1] and data[index + 1][4] == "GRANTED" and "OUT" in data[index + 1][3] \
                        and start.date() == data[index + 1][2].date():
                    finish = data[index + 1][2]
                    duration = finish - start
                    try:
                        if int(entry[1] * 1) > 1:
                            tgen_total += duration
                    except:
                        if "CON" in entry[1]:
                            con_total += duration
                        else:
                            total += duration

                    print(entry, "\n", data[index + 1])
                    print("Tgen: ", tgen_total)
                    print("Con: ", con_total)
                    print("Other: ", total)

            index += 1

        except IndexError:
            break
