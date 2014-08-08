#!/usr/bin/python


import xlrd
import xlwt
import openpyxl

import sys
import re

class SpreadSheetQuickSheet97(object):
    def __init__(self,h):
        self.h = h

    def name(self):
        return self.h.name
    
    def row_range(self):
        return (0,self.h.nrows - 1)
    
    def row_count(self):
        return self.h.nrows
    
    def col_range(self):
        return (0,self.h.ncols - 1)
 
    def col_count(self):
        return self.h.ncols
    
    def cell_value(self,row,col):
        return self.h.cell_value(row,col)
    
    def sync(self,old_xworksheet,new_xworksheet):
        pass
    

        

class SpreadSheetQuickSheet07(object):
    def __init__(self,h):
        self.h = h

    def name(self):
        return self.h.title
    
    def row_range(self):
        return (0,len(self.h.rows) - 1)
    
    def row_count(self):
        return self.h.get_highest_row()
    
    def col_range(self):
        return (0,len(self.h.columns) - 1)

    def col_count(self):
        if self.h.get_highest_column() > 256:
            return 256
        return self.h.get_highest_column()
    
    def cell_value(self,xrow,xcol):
        
        cell = self.h.cell(row = xrow + 1,column = xcol + 1)
        if cell.value is None:
            return ''
        try:
            s = str(cell.value)
        except:
            s = cell.value
        return s
    
class SpreadSheetQuickSheet(object):
    def __init__(self,h = None):
        self.h = h

    def name(self):
        return ''
    
    def row_range(self):
        return (0,0)
    
    def row_count(self):
        return 0
    
    def col_range(self):
        return (0,0)

    def col_count(self):
        return 0
    
    def cell_value(self,xrow,xcol):
        return ''
    
class SpreadSheetQuick(object):
    def __init__(self,fname):
        self.fname = fname
        if re.match(r'.*\.xlsx$',self.fname.lower()):
            self.fmt = '2007'
            self.workbook = openpyxl.load_workbook(filename = fname)
        elif re.match(r'.*\.xls$',self.fname.lower()):
            self.fmt = '1997'
            self.workbook = xlrd.open_workbook(fname)
        else:
            self.fmt = ''
            
    def worksheets(self):
        if self.fmt == '2007':
            return self.workbook.get_sheet_names()
        elif self.fmt == '1997':
            return self.workbook.sheet_names()
    
    def worksheet(self,name):
        if self.fmt == '2007':
            return SpreadSheetQuickSheet07(self.workbook.get_sheet_by_name(name))
        elif self.fmt == '1997':
            return SpreadSheetQuickSheet97(self.workbook.sheet_by_name(name))       

    @staticmethod
    def create(xworkbook,filename):
        if re.match(r'.*\.xls$',filename.lower()):
            workbook = xlwt.Workbook()
            for xworksheet in xworkbook:
                sheet = workbook.add_sheet(xworksheet['name'])
                for d in xworksheet['data'].keys():
                    v = xworksheet['data'][d]
                    index = str(d).split(',')
                    sheet.write(int(index[0]),int(index[1]),v)
            try:
                workbook.save(filename)
                return True
            except:
                return False
            
        return False
    
    @staticmethod
    def save(xworkbook,filename):
        if re.match(r'.*\.xls$',filename.lower()):
            return SpreadSheetQuick.create(xworkbook,filename)
            
        return False    

def XlsHeader(i):
    if i >= 0 and i <= 25:
        return chr(ord('A') + i)
    else:
        t = int(i / 26)
        return chr(ord('A') + t - 1) + chr(ord('A') + i - t * 26)
