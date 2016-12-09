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
#------------------------------------------------
# ----- final_result is a dict like this:       |
# ----- {'keyword_1': [                         |
# -----                {index_1: 'feedback_1'}, |
# -----                {index_2: 'feedback_2'}  |
# -----               ],                        |
# -----  'keyword_2': [{...},                   |
# -----                {...}                    |
# -----               ]                         |
# ----- }                                       |
#------------------------------------------------


def help():
    print \
'''
Requirement:
      Firstly, you should install xlrd and xlwt module to
      read&write excel, use the command below:

      on Ubuntu(maybe need sudo privilege):
      pip install xlrd
      pip install xlwt

      on Windows:
      python -m pip install xlrd
      python -m pip install xlwt

Usage:
      python main.py source_file "keyword1|keyword2"

Example:
      python main.py test_check.xlsx "233|亮|暗"
'''


def do_check(keywordlist):
    print 'Now we check'
    # get the keyword, and init the final result structure
    for item in keywordlist:
        final_result[item] = []
    # open excel
    check_file = xlrd.open_workbook(filename)
    # get the 1st sheet page, can modify it
    opened_sheet = check_file.sheet_by_index(0)
    result = []
    for row_index in range(opened_sheet.nrows):
        # get the 7th column cell value, can modify it
        unchecked_item = opened_sheet.cell_value(rowx=row_index, colx=6)
        # change cell value to string or unicode to make it easy to store and compare
        if type(unchecked_item) is not unicode:
            result.append(str(unchecked_item))
        else:
            result.append(u''.join(unchecked_item).encode('utf-8').strip())
    # define the index to mark the position of item we've found in the whole sheet
    final_index = 0
    # do the search
    for item in result:
        if(type(item) is not unicode):
            for keyword_item in keywordlist:
                if(item.decode('utf-8').find(keyword_item.decode('utf-8')) != -1):
                    single_item = {final_index: item}
                    final_result[keyword_item].append(single_item)

        final_index = final_index + 1
    # output the result to terminal
    for x in range(0, len(keywordlist)):
        print 'keyword is %s:' % keywordlist[x]
        if len(final_result[keywordlist[x]]) == 0:
            print 'Empty'
            continue
        for item in final_result[keywordlist[x]]:
            for key, value in item.iteritems():
                print 'index is %s, feedback is %s' % (key, value)

if __name__ == "__main__":
    if (len(sys.argv) < 2):
        help()
        exit()
    else:
        # get the filename and keyword in command line parameters
        filename = sys.argv[1]
        raw_keywordlist = sys.argv[2].rstrip().split("|")
        if (not os.path.exists(filename)):
            print 'Cann\'t find the file --- %s' % filename
            exit()
        do_check(raw_keywordlist)
