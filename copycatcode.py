#!/usr/bin/env python

import urllib
import json
import sys
import xlrd


from xlrd import open_workbook

#print open_workbook('EOQ.xls')

book = open_workbook('ETK_test.xls')
sheet = book.sheet_by_index(0)
#print sheet.name
#print sheet.nrows
#print sheet.ncols
for row_index in range(sheet.nrows): #prints every row in a specific column
    #for col_index in range(sheet.ncols):
        #print cellname(row_index,col_index),'-',
    print sheet.cell(row_index,1).value #prints just column 1 rn

for row_index in range(sheet.nrows):
    for row_index in range(sheet.nrows):
        starting_point = sheet.cell(row_index,1).value
        end_point = sheet.cell(row_index,1).value

        googleaddress = "http://maps.googleapis.com/maps/api/directions/json?"
 
        response = urllib.urlopen(googleaddress + "origin=" + starting_point + "&destination=" + end_point + "&sensor=false&mode=driving")

        pyresponse = json.load(response)

        results = pyresponse["routes"]

        for i in range(len(results)):     

             for key in results[i]:

                  if key =="legs":

                       print results[i][key][0]["distance"]["text"]
                       print results[i][key][0]["duration"]["text"]



  











