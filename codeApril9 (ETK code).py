#!/usr/bin/env python

import urllib
#import urllib2
import json
import sys
import xlrd
import xlwt


from xlrd import open_workbook
from xlwt import Workbook
from xlwt import Worksheet
from xlwt import easyxf


class MyException(Exception):
    pass

#print open_workbook('EOQ.xls')
save_file_as = 'Output.xls'

book = open_workbook('ETK_test.xls')
sheet = book.sheet_by_index(0)
#print sheet.name
#print sheet.nrows
#print sheet.ncols
#for row_index in range(sheet.nrows): #prints every row in a specific column
    #for col_index in range(sheet.ncols):
        #print cellname(row_index,col_index),'-',
    #print sheet.cell(row_index,1).value #prints just column 1 rn

thisCount = 0;
Matrix_Time = [[0 for x in range(sheet.nrows-1)] for x in range(sheet.nrows-1)]
Matrix_Distance = [[0 for x in range(sheet.nrows-1)] for x in range(sheet.nrows-1)]
Matrix_starts = [0 for x in range(sheet.nrows-1)]
Matrix_ends = [0 for x in range(sheet.nrows-1)]
for row_indexA in range(1, sheet.nrows):
    for row_indexB in range(1, sheet.nrows):
        #print "Current Row Index A: "
        #print row_indexA
        #print "Current Row Index B: "
        #print row_indexB
        thisCount = thisCount +1
        #print(thisCount)
        starting_point = sheet.cell(row_indexA,2).value
        end_point = sheet.cell(row_indexB,4).value
        Matrix_starts[row_indexA-1] = starting_point #stores start_points
        Matrix_ends[row_indexB-1] = end_point #stores end_points
        print "Starting Point: "
        print starting_point
        print "End Point:  "
        print end_point

        googleaddress = "http://maps.googleapis.com/maps/api/directions/json?"
        
 #       try:
        response = urllib.urlopen(googleaddress + "origin=" + starting_point + "&destination=" + end_point + "&sensor=false&mode=driving")
 #       except urllib2.URLError, e:
 #           raise MyException("There was an error: %r" % e)
        
        pyresponse = json.load(response)

        results = pyresponse["routes"]
        print len(results);

        for i in range(len(results)):     

             for key in results[i]:

                  if key =="legs":
                        Matrix_Time[row_indexA-1][row_indexB-1] = results[i][key][0]["duration"]["text"]
                        print "Distance from " + starting_point + " to " + end_point + " " + Matrix_Time[row_indexA-1][row_indexB-1]
                        Matrix_Distance[row_indexA-1][row_indexB-1] = results[i][key][0]["distance"]["text"]
                        print "Distance from " + starting_point + " to " + end_point + " " + Matrix_Distance[row_indexA-1][row_indexB-1]

                        print results[i][key][0]["distance"]["text"]
                        print results[i][key][0]["duration"]["text"]
answer_workBook = Workbook(encoding='ascii',style_compression=0)
#distance_sheet = Worksheet("distance_sheet",answer_workBook)
the_style = easyxf()
#print len(Matrix_Distance)
#print len(Matrix_Distance) - 1
count = 0
distance_sheet = answer_workBook.add_sheet('distance_sheet')
distance_sheet.write(0,0,'v starting points v => ending points =>')
for x in range(0, row_indexA):
    distance_sheet.write(x+1,0,Matrix_starts[x])
for x in range(0, row_indexB):
    distance_sheet.write(0,x+1,Matrix_ends[x])
for x in range(1, (len(Matrix_Distance)+1)):
    for y in range(1, (len(Matrix_Distance)+1)):
        print('new iteration')
        count = count+1;
        print(count)
        distance = Matrix_Distance[x-1][y-1]
        print(distance);
        print(x)
        print(y)
        distance_sheet.write(x,y,distance)

answer_workBook.save(save_file_as) 