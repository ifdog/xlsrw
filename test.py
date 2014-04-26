#coding:utf8
__author__ = 'ifdog'
__version__ = 0.1

import  xlrw
f = 'c:\\V01.xls'
a = xlrw.Excel(f)

print a[1][1][2]
for i in a:
    for j in i:
        for k in j:
            print k,
        print