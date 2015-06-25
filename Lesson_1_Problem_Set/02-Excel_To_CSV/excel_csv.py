# -*- coding: utf-8 -*-
# Find the time and value of max load for each of the regions
# COAST, EAST, FAR_WEST, NORTH, NORTH_C, SOUTHERN, SOUTH_C, WEST
# and write the result out in a csv file, using pipe character | as the delimiter.
# An example output can be seen in the "example.csv" file.
import xlrd
import os
import csv
from zipfile import ZipFile
datafile = "2013_ERCOT_Hourly_Load_Data.xls"
outfile = "2013_Max_Loads.csv"


def open_zip(datafile):
    with ZipFile('{0}.zip'.format(datafile), 'r') as myzip:
        myzip.extractall()


def parse_file(datafile):
    workbook = xlrd.open_workbook(datafile)
    sheet = workbook.sheet_by_index(0)
    data = []

    for i in range(sheet.ncols - 2):
        name = sheet.cell_value(0, 1 + i).strip('u')
        col = sheet.col_values(i + 1, start_rowx = 1, end_rowx = None)
        max_val = max(col)
        max_pos = col.index(max_val) + 1
        maxtime = sheet.cell_value(max_pos, 0)
        rtime = xlrd.xldate_as_tuple(maxtime, 0)


        entry = [name, rtime[0], rtime[1], rtime[2], rtime[3], max_val]

        data.append(entry)

    # YOUR CODE HERE
    # Remember that you can use xlrd.xldate_as_tuple(sometime, 0) to convert
    # Excel date to Python tuple of (year, month, day, hour, minute, second)
    return data

def save_file(data, filename):

    header = ['Station', 'Year', 'Month', 'Day',
        'Hour', 'Max Load']

    with open(filename, 'wb') as csvfile:
        csvwrite = csv.writer(csvfile, delimiter='|')
        csvwrite.writerow(header)

        for elem in data:
            csvwrite.writerow(elem)

    
def test():
    #open_zip(datafile)
    data = parse_file(datafile)
    save_file(data, outfile)

    ans = {'FAR_WEST': {'Max Load': "2281.2722140000024", 'Year': "2013", "Month": "6", "Day": "26", "Hour": "17"}}
    
    fields = ["Year", "Month", "Day", "Hour", "Max Load"]
    with open(outfile) as of:
        csvfile = csv.DictReader(of, delimiter="|")
        for line in csvfile:
            s = line["Station"]
            if s == 'FAR_WEST':
                for field in fields:
                    assert ans[s][field] == line[field]

        
test()