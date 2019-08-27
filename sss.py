


with open('Door HistoryJulyWPS.csv') as fh:
    # Skip initial comments that starts with #
    while True:
        line = fh.readline()
        # break while statement if it is not a comment line
        # i.e. does not startwith #
        if not line.startswith('CardHolder'):
            break


    # Second while loop to process the rest of the file
    while line:
        print(line)