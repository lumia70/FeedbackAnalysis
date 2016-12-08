#! /usr/bin/env python2
# -*- coding=utf-8 -*-

import re
import os
import sys
import xlrd
import xlwt
from datetime import datetime

filename = ""
final_result = {}

def help():
    print \
'''
Usage:
      python main.py source_file "keyword1|keyword2"
'''

def doCheck(keywordlist):
    print 'Now we check'
    for item in keywordlist:
        final_result[item] = []
    #print final_result
    #exit()
    check_file = xlrd.open_workbook(filename)
    opened_sheet = check_file.sheet_by_index(0)
    #print opened_sheet.cell_value(rowx=18, colx=6)
    result = []
    for row_index in range(opened_sheet.nrows):
        #print 'now at index ' + str(row_index)
        #print type(opened_sheet.cell_value(rowx = row_index, colx = 6)) is not unicode
        if type(opened_sheet.cell_value(rowx = row_index, colx = 6)) is not unicode:
            #print 'not unicode'
            result.append(str(opened_sheet.cell_value(rowx = row_index, colx = 6)))
        else:
            #print 'unicode'
            result.append(u''.join(opened_sheet.cell_value(rowx = row_index, colx = 6)).encode('utf-8').strip())
        #result.append(u''.join(opened_sheet.cell_value(rowx = row_index, colx = 6)).encode('utf-8').strip())
        #if ((u''.join(opened_sheet.cell_value(rowx = row_index, colx = 6)).encode('utf-8').strip()).find(u'亮') != -1):
        #    print 'find it! row index is ' + str(row_index) + ' ' + opened_sheet.cell_value(rowx = row_index, colx = 6)
    finalIndex = 0
    for item in result:
        #print type(item)
        #print item
        if(type(item) is not unicode):
            for keywordItem in keywordlist:
                if(item.decode('utf-8').find(keywordItem.decode('utf-8')) != -1):
                    #print item
                    #print keywordItem
                    singleItem = {finalIndex:item}
                    final_result[keywordItem].append(singleItem)
                #else:
                #    print 'not found %s' %keywordItem

        finalIndex = finalIndex + 1
    #list(map(print, final_result[keywordlist[2]])) #not support in python 2.x
    #print final_result[keywordlist[2]]
    for item in final_result[keywordlist[2]]:
        #print item
        for key, value in item.iteritems():
            print 'index is %s, feedback is %s' % (key, value)

if __name__ == "__main__":
    if (len(sys.argv) < 2):
        help()
        exit()
    else:
        filename = sys.argv[1]
        raw_keywordlist = sys.argv[2].rstrip().split("|")
        #print raw_keywordlist[2].decode('utf-8').encode('utf-8')
        if (not os.path.exists(filename)):
            print 'Cann\'t find the file --- %s' %filename
            exit()
        doCheck(raw_keywordlist)
