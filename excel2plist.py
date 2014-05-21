from xlrd import open_workbook, cellname
import os

excel_file_name = 'iPhone%26Android_Weibo_Skin_Configs.xlsx'
wb = open_workbook(excel_file_name)

sheet = wb.sheet_by_index(0)#New for 4.1.0
plist_head='''<?xml version="1.0" encoding="UTF-8"?><!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd"><plist version="1.0"><dict>'''
plist_tail='''</dict></plist>'''
values = []
values.append(plist_head)

for row_index in range(sheet.nrows):

  #skip col title
  if row_index == 0:
    continue

  pKey = sheet.cell(row_index,1).value #key for iPhone col
  pValue = sheet.cell(row_index,4).value #Default col

  #skip blank line
  if pKey == '':
    continue

  values.append("<key>")
  values.append(pKey)
  values.append("</key>")

  if sheet.cell(row_index,0).value == 'bool':
    if pValue == 'YES':
      values.append("<true/>")
    else:
      values.append("<false/>")
  elif sheet.cell(row_index,0).value == 'digit':
    values.append("<integer>")
    values.append(pValue)
    values.append("</integer>")
  else:
    values.append("<string>")
    values.append(pValue)
    values.append("</string>")

#append plist tail
values.append(plist_tail)

#convert to string
str_convert = ''.join(values)

#create new plist file
target = open ('config.plist', 'w+')
target.write(str_convert)

print "done"
target.close()

#remove excel file
os.remove(excel_file_name)
