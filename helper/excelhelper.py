import xlrd
import xlwt
import os
from xlutils.copy import copy

row_default_style = xlwt.Style.easyxf(
    'pattern:fore_color white; font: color black,height 200; align: horiz center,vert center')
row_default_style.alignment.wrap = xlwt.Alignment.WRAP_AT_RIGHT


head_row_default_style = xlwt.Style.easyxf(
    'pattern: pattern solid, fore_color blue;font: color white,bold on,height 300; align: horiz center,vert center')

default_row_height = 256*2
default_row_width = 256*30


def getRowCount(filepath, sheet_index):
    wb = xlrd.open_workbook(filepath)
    sheet = wb.sheet_by_index(sheet_index)
    return sheet.nrows


def getColCount(filepath, sheet_index):
    wb = xlrd.open_workbook(filepath)
    sheet = wb.sheet_by_index(sheet_index)
    return sheet.ncols


def getSheet(filepath, sheet_index):
    wb = xlrd.open_workbook(filepath)
    sheet = wb.sheet_by_index(sheet_index)
    return sheet


def getColValues(filepath, sheet_index, col):
    col -= 1
    wb = xlrd.open_workbook(filepath)
    sheet = wb.sheet_by_index(sheet_index)
    values = []

    for row in range(sheet.nrows):
        for rowCol in range(sheet.ncols):

            if rowCol == col:
                cel = sheet.cell(row, rowCol)
                val = cel.value
                if val.isalpha():
                    values.append(val)
                else:
                    values.append(str(val))

    return values


def inserData(sheet, data, insertStartIndex, row_height=default_row_height, col_width=default_row_width, head_style=head_row_default_style, other_style=row_default_style):
    for i in range(0, len(data)):
        rowData = data[i]
        sheet.row(i+insertStartIndex).height_mismatch = True
        sheet.row(i+insertStartIndex).height = row_height

        for j in range(0, len(rowData)):
            sheet.col(j).width = col_width

            if head_style != None and i == 0 and insertStartIndex == 0:
                sheet.write(i+insertStartIndex, j, rowData[j], head_style)
            else:
                sheet.write(i+insertStartIndex, j, rowData[j], other_style)


def newExcelInsers(filepath, sheetname, head, data, row_height=default_row_height, col_width=default_row_width, head_style=head_row_default_style, other_style=row_default_style):
    if os.path.exists(filepath):
        os.remove(filepath)


    dirname=os.path.dirname(filepath)
    if not  os.path.exists(dirname):
        os.makedirs(dirname)
    
    newWb = xlwt.Workbook()
    sheet = newWb.add_sheet(sheetname, cell_overwrite_ok=True)

    hs = None
    if(len(head) > 0):
        data.insert(0, head)
        if head_style == None:
            hs = head_row_default_style
        else:
            hs = head_style
    inserData(sheet, data, 0, row_height=row_height,
              col_width=col_width, head_style=hs, other_style=other_style)
    newWb.save(filepath)


def appendExcel(file_path, sheet_index, data, row_height=default_row_height, col_width=default_row_width, other_style=row_default_style):
    rb = xlrd.open_workbook(file_path)
    rbSheet = rb.sheet_by_index(sheet_index)
    startIndexRow = rbSheet.nrows
    wb = copy(rb)

    rWsheet = wb.get_sheet(sheet_index)
    inserData(rWsheet, data, startIndexRow, row_height=row_height,
              col_width=col_width, other_style=other_style, head_style=None)
    wb.save(file_path)


