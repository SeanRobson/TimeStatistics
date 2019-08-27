# import module to read and write in .xlsx files
import pandas as pd
# import module to dates and time
import datetime as dt
import datetime as datetime

def load_data():
    """
    Open the excel file, get number of rows and columns
    """

    # open the excel data file
    #input_file = pd.read_excel('lil processy.xlsx')
    input_file = pd.read_csv('Door HistoryJulyWPS.csv')
    # determine the size of the excel file (important to get number of rows)
    row, column = input_file.shape

    # open CSV file
    #input_file = pd.read_csv("Door HistoryJulyWPS.csv")
    # smartly read the file

    # get number of rows

    """
    Read the excel file, get the important information, change to correct data type if necessary
    """

    # create data storage array
    data = []

    # denied counter
    denied = 0

    # iterate over each entry and store important information
    for entry in range(0, row):

        # create temp data storage array (empty each iteration)
        temp = []

        # get information
        name = input_file.iloc[entry, 0]
        ID = input_file.iloc[entry, 2]
        time = pd.to_datetime(input_file.iloc[entry, 4])
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

    return data, row, denied


data, row, denied = load_data()
data.sort(key=lambda x:(x[0], x[2]))


index = 0
denied = 0

tgen_total = dt.timedelta(0,0,0)
con_total = dt.timedelta(0,0,0)
total = dt.timedelta(0,0,0)


for entry in data:
    try:
        if entry[4] != "GRANTED":
            denied += 1
            # capture if the same person
        elif "IN" in entry[3] and entry[4] == "GRANTED":
            start = entry[2]
            ID = entry[1]
            if ID == data[index+1][1] and data[index+1][4] == "GRANTED" and "OUT" in data[index+1][3] \
                    and start.date() == data[index+1][2].date():
                finish = data[index+1][2]
                duration = finish - start
                try:
                    if int(entry[1]*1)>1:
                        tgen_total += duration
                except:
                    if "CON" in entry[1]:
                        con_total += duration
                    else:
                        total += duration

                print(entry, "\n", data[index+1])
                print("Tgen: ", tgen_total)
                print("Con: ", con_total)
                print("Other: ", total)

        index += 1

    except IndexError:
        break
