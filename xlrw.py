#coding:utf8
__author__ = 'ifdog'
__version__ = 0.1

import xlrd
import xlwt
import xlutils

#################################################################################################
## 对xlrd,xlwt,xlutils等包重新封装,实现形如对xls文件以形式 xlsfile['表名']['行号']['列号'] 访问的类库.
## A library based in the modules xlrd,xlwt,xlutils for providing another way to access xls files
## with the following looking: value = xlsfile['sheet']['row']['col']
#################################################################################################


class Excel(object):
    def __init__(self, filename):
        self.path = filename
        try:
            self.file = xlrd.open_workbook(self.path)
        except Exception as e:  # TODO:自定义异常类
            print e
        self.sheets = self.file.sheets()  # 表单对象
        self.sheet_names = self.file.sheet_names()  # 表单名
        self.sheet_num = self.file.nsheets
        self.current_sheet = None

    def __getitem__(self, item):  # 模拟容器方式取工作表
        if isinstance(item, int):
            try:
                s = Sheet(self.file)
                return s.sheet_by_index(item)
            except:
                raise IndexError(u'Index Error')
        elif isinstance(item, (str, unicode)):  # 此处是坑
            try:
                s = Sheet(self.file)
                return s.sheet_by_name(item)
            except:
                raise KeyError(u'Key Error')
        else:
            raise TypeError(u'key type Error')

    def sheet(self, num=None, name=None):  # 调用方法取工作表
        if name:
            self.current_sheet = self.file.sheet_by_name(name)
        else:
            self.current_sheet = self.file.sheet_by_index(num)
        return self.current_sheet


class Sheet(object):
    # TODO:col字母标识到index的转换
    def __init__(self, file):
        self.data = None
        self.file = file
        self.rows = None
        self.cols = None
        self.sh = None

    def sheet_by_name(self, name):
        self.sh = self.file.sheet_by_name(name)
        self.rows = self.sh.nrows
        self.cols = self.sh.ncols
        return self

    def sheet_by_index(self, index):
        self.sh = self.file.sheet_by_index(index)
        self.rows = self.sh.nrows
        self.cols = self.sh.ncols
        return self

    def __getitem__(self, item):
        if isinstance(item, int):
            if -1 < item < self.rows:
                return self.sh.row_values(item)
            else:
                raise IndexError('Index Value Error')
        else:
            raise TypeError('Key Type Error')

    def get_row(self, row):
        return self.sh.row_values(row)

    def get_col(self, col):
        return self.sh.col_values(col)

    def get_cell(self, row, col):
        return self.sh.cell(rowx=row, colx=col).value