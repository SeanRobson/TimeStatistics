"""

This code was written by Sean "Charles" Robson for Territory Generation. Any unauthorised distribution of this code
without consent from the  developer will be seen as a breach of copy right and intellectual property laws.
Please seek consent from Charles on sean.robson@territorygeneration.com.au or call 0467694844.

This code is designed to:
    - read data from .xlsx file
    -

"""

# import module to read and write in .xlsx files
import pandas as pd
# import module to handle dates and time from .xlsx file
from datetime import datetime

input_file = pd.read_excel('JanuaryData.xlsx')

# determine the size of the excel file (important to get number of rows)
row, column = input_file.shape

# empty array to hold the data
RawData = []
count = 1
VisitorList = {}

# go through the cells and append data to the RawData array as a list of lists
for rows in range(0, row):
    # clear the temporary data holder
    DataHolder = []

    # get data
    SurnameRaw = input_file.iloc[rows, 0]
    FirstNameRaw = input_file.iloc[rows, 1]
    AGSRaw = str(input_file.iloc[rows, 2])
    if AGSRaw == "nan":

        if SurnameRaw in VisitorList.keys():
            AGSRaw = VisitorList[SurnameRaw]

        else:
            AGSRaw = str('visitor' + str(count))
            VisitorList[SurnameRaw] = AGSRaw
            count += 1

    StatusRaw = input_file.iloc[rows, 3]
    TimeRaw = str(input_file.iloc[rows, 4])
    DateRaw = str(datetime.date(input_file.iloc[rows, 5]))
    LocationRaw = input_file.iloc[rows, 6]

    # split up the raw date and time date to make one container in the format for processing
    TimeData = TimeRaw.split(":")
    hour = int(TimeData[0])
    minute = int(TimeData[1])
    DateData = DateRaw.split("-")
    year = int(DateData[0])
    month = int(DateData[1])
    day = int(DateData[2])

    # store the date and time in one variable
    DateTime = datetime(year, month, day, hour, minute)

    # make a list for the data
    DataHolder = [SurnameRaw, FirstNameRaw, AGSRaw, StatusRaw, DateTime, LocationRaw]

    # append to the RawData array (list of lists)
    RawData.append(DataHolder)

""" 

indexing in python is [1D][2D][3D][4D]...[nD]

Indexing for and corresponding data type for RawData
0 = SURNAME
1 = FIRST NAME
2 = AGS
3 = STATUS
4 = DATE TIME
5 = LOCATION

"""
CIPStime = 0
CIPSCONtime = 0
CIPSVIStime = 0
CIPSTOTALtime = 0
WPStime = 0
WPSCONtime = 0
WPSVIStime = 0
WPSTOTALtime = 0
KPStime = 0
KPSCONtime = 0
KPSVIStime = 0
KPSTOTALtime = 0

# iterating over every row
for entry in range(0, row):
    # saving the entry number to row number as a placeholder to be manipulated
    RowNum = entry

    # checking if the status is in (entering)
    if RawData[entry][3] == 'In':
        # AGS of the person entering
        SearchAGS = RawData[entry][2]
        # time of the person entering
        TimeStart = RawData[entry][4]

        # iterating over every row that is after this time
        for SubsequentRows in range(RowNum, row):

            # checking if next rows have the same AGS and if the status is out (leaving)
            if RawData[SubsequentRows][2] == SearchAGS and RawData[SubsequentRows][3] == 'Out':
                # capture time out (exit)
                TimeOut = RawData[SubsequentRows][4]

                # calculate time difference
                Duration = str(TimeOut - TimeStart)
                HoursDuration = int(Duration.split(":")[0])
                MinutesDuration = int(Duration.split(":")[1])

                # increment total times for contractor, staff and visitors
                if RawData[entry][5] == "CIPS" and RawData[entry][2][:3] == "CON":
                    # increment CIPS total time
                    CIPSCONtime += HoursDuration + (MinutesDuration / 60)
                elif RawData[entry][5] == "CIPS" and RawData[entry][2][:3] == "vis":
                    # increment CIPS total time
                    CIPSVIStime += HoursDuration + (MinutesDuration / 60)
                elif RawData[entry][5] == "CIPS" and RawData[entry][2][:3] != "CON" and RawData[entry][2][:3] != "vis":
                    # increment CIPS total time
                    CIPStime += HoursDuration + (MinutesDuration / 60)
                elif RawData[SubsequentRows][5] == "WPS" and RawData[SubsequentRows][2][:3] == "CON":
                    # increment CIPS total time
                    WPSCONtime += HoursDuration + (MinutesDuration / 60)
                elif RawData[SubsequentRows][5] == "WPS" and RawData[SubsequentRows][2][:3] == "vis":
                    # increment CIPS total time
                    WPSVIStime += HoursDuration + (MinutesDuration / 60)
                elif RawData[SubsequentRows][5] == "WPS" and RawData[entry][2][:3] != "CON" and RawData[entry][2][:3] != "vis":
                    # increment CIPS total time
                    WPStime += HoursDuration + (MinutesDuration / 60)
                elif RawData[SubsequentRows][5] == "KPS" and RawData[SubsequentRows][2][:3] == "CON":
                    # increment CIPS total time
                    KPSCONtime += HoursDuration + (MinutesDuration / 60)
                elif RawData[SubsequentRows][5] == "KPS" and RawData[SubsequentRows][2][:3] == "vis":
                    # increment CIPS total time
                    KPSVIStime += HoursDuration + (MinutesDuration / 60)
                elif RawData[SubsequentRows][5] == "KPS" and RawData[entry][2][:3] != "CON" and RawData[entry][2][:3] != "vis":
                    # increment CIPS total time
                    KPStime += HoursDuration + (MinutesDuration / 60)

# total times
CIPSTOTALtime = CIPSCONtime + CIPSVIStime + CIPStime
WPSTOTALtime = WPStime + WPSCONtime + WPSVIStime
KPSTOTALtime = KPStime + KPSCONtime + KPSVIStime

# print total times for each location (site)
print("CIPS CON : ", CIPSCONtime)
print("CIPS VIS : ", CIPSVIStime)
print("CIPS STF : ", CIPStime)
print("CIPS TOT : ", CIPSTOTALtime)
print("WPS CON : ", WPSCONtime)
print("WPS VIS : ", WPSVIStime)
print("WPS STF : ", WPStime)
print("WPS TOT : ", WPSTOTALtime)
print("KPS CON : ", KPSCONtime)
print("KPS VIS : ", KPSVIStime)
print("KPS STF : ", KPStime)
print("KPS TOT : ", KPSTOTALtime)