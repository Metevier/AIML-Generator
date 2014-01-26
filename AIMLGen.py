#Written by Tyler Metevier 8/20/13
#Version 1.2
import xlrd
import sys
import os

#sheet, row and file
def genAIML(sheet,row,f):

    #Find current section's end point, end_row
    end_row = 0
    for rx in range(row+1,sheet.nrows):
        if sheet.row(rx)[0].value == "Pattern":
            end_row = rx
            break
    if end_row == 0:
        end_row = sheet.nrows

    #Find initial pattern, initial_pattern
    initial_pattern = sheet.row(row)[1].value

    #Find start of search strings, string_start
    template = []
    for rx in range(row+1,end_row):
        if sheet.row(rx)[0].value == "Search Strings":
            string_start = rx
            break
        else:
            template.append(sheet.row(rx)[1].value)
    
    #Generate search strings array, strings
    strings = []
    for rx in range(string_start,end_row):
        strings.append(sheet.row(rx)[1].value)

    #Generate display array, steps
    steps = []
    steps.append(template[0] + '\n\t<ul>')
    for t in range(1,len(template)):
        steps.append('\t\t<li>' + template[t] + '</li>')
    steps.append('\t</ul>\n')

    #Write data to file
    writeFile(f,initial_pattern,'\n'.join(steps))
    for st in strings:
        writeFile(f,st,'<srai>' + initial_pattern + '</srai>')
    
    #Recurse if not file end
    if end_row != sheet.nrows:
        genAIML(sheet,end_row,f)

#Write pattern and template to file
def writeFile(f,pattern,template):
    f.write('<category>\n\t<pattern>' + pattern + '</pattern>\n')
    f.write('\t<template>' + template + '</template>\n</category>\n')


#Start Main
#Check for excel workbook

file_exists = False
while file_exists == False:
    sys.stdout.write("Enter in Workbook name: ")
    inFile = raw_input()
    file_exists = os.path.isfile(inFile)
    if file_exists == False:
        print "File not found."
    else:
        print "Successfully opened file!"

#Open Workbook
book = xlrd.open_workbook(inFile)

#Get and Open Worksheet
sys.stdout.write( "Please Enter a Worksheet name: ")
sheet_input = raw_input()
sheet = book.sheet_by_name(sheet_input)

#Create out file
sys.stdout.write( "What would you like to name the file? ")
f_name = raw_input()
f = open(f_name,'w+')

#Find Starting Point
for rx in range(sheet.nrows):
    if sheet.row(rx)[0].value == "Pattern":
        row = rx
        break

#Write Opening XML
f.write('<?xml version=\"1.0\" encoding=\"ISO-8859-15\"?>\n<aiml>\n')

#Generate AIML
genAIML(sheet,row,f)

#Write Closing XML
f.write('\n</aiml>')

#File Close
f.close()