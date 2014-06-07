#!/usr/bin/env python

from xlrd import open_workbook, cellname
import os

excel_file_name = os.path.dirname(os.path.realpath(__file__)) + '/iPhone%26Android_Weibo_Skin_Configs.xlsx'
wb = open_workbook(excel_file_name)

sheet = wb.sheet_by_index(0)#New for 4.1.0
plist_head='''<?xml version="1.0" encoding="UTF-8"?><!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd"><plist version="1.0"><dict>'''
plist_tail='''</dict></plist>'''
default_value_dic = {}


def generatePlist(convert_col_num, is_default = False):

  plist_name = ''
  values = []
  values.append(plist_head)

  for row_index in range(sheet.nrows):

    #skip col title
    if row_index == 0:
      plist_name = sheet.cell(row_index,convert_col_num).value
      plist_name = plist_name.encode('utf-8') +'.plist'
      continue

    pKey = sheet.cell(row_index,1).value #key for iPhone col
    pValue = sheet.cell(row_index,convert_col_num).value #convert col

    #skip blank line
    if pKey == '':
      continue

    values.append("<key>")
    values.append(pKey)
    values.append("</key>")

    if pValue == '' and is_default == False:
      pValue = default_value_dic[pKey]

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

    #save default value dic
    if is_default == True:
      default_value_dic[pKey] = pValue

  #append plist tail
  values.append(plist_tail)



  #convert to string
  str_convert = ''.join(values).encode('utf-8')

  directory = os.path.dirname(os.path.realpath(__file__)) + '/plists/'
  if not os.path.exists(directory):
      os.makedirs(directory)

  #create new plist file
  target = open (directory+plist_name, 'w+')
  target.write(str_convert)

  print plist_name+' done'
  target.close()

for i in range(4, sheet.ncols):
  if i == 4:
    generatePlist(i, True)
  else:
    generatePlist(i)

#remove excel file
#os.remove(excel_file_name)
