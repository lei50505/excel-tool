#! /usr/bin/env python
# -*- coding: utf-8 -*-

import sys
sys.path.append(".")

from lib.utils import *

inputExcelPath = "input.xlsx"
sheet1NumberColumn = 0
sheet2NumberColumn = 0
sheet1MaxColumn = 0
sheet2MaxColumn = 0
sheet1MaxRow = 0
sheet2MaxRow = 0

def main():
    return

def initInputExcel():
    removeFile(inputExcelPath)
    book = createBook()
    sheet = activeBook(book)
    setSheetName(sheet, "Sheet1")
    addSheet(book, "Sheet2")
    saveBook(book, inputExcelPath)
    return

def checkInputExcel():

    global sheet1NumberColumn
    global sheet2NumberColumn
    global sheet1MaxColumn
    global sheet2MaxColumn
    global sheet1MaxRow
    global sheet2MaxRow
    
    book = loadBook("a.xlsx")
    
    sheet1 = getSheetByName(book,"Sheet1")
    sheet1MaxColumn = getMaxColumn(sheet1)
    sheet1MaxRow = getMaxRow(sheet1)
    
    sheet2 = getSheetByName(book,"Sheet2")
    sheet2MaxColumn = getMaxColumn(sheet2)
    sheet2MaxRow = getMaxRow(sheet2)

    flag = False
    for column in range(1,sheet1MaxColumn+1):
        cell = getCell(sheet1,1,column)
        if not cell.value:
            continue
        if cell.font.b == True and cell.font.i == True:
            sheet1NumberColumn = column
            flag = True

    for column in range(1,sheet2MaxColumn+1):
        cell = getCell(sheet2,1,column)
        
        if not cell.value:
            continue
        
        
        if cell.font.b == True and cell.font.i == True:
            sheet2NumberColumn = column
            flag = True

    if flag == False:
        return
    
    return

findNum = {}


def procData():
    global sheet1NumberColumn
    global sheet2NumberColumn
    global sheet1MaxColumn
    global sheet2MaxColumn
    global sheet1MaxRow
    global sheet2MaxRow
    
    book = loadBook("a.xlsx")
    
    sheet1 = getSheetByName(book,"Sheet1")
    sheet2 = getSheetByName(book,"Sheet2")
    
    for sheet1Row in range(1,sheet1MaxRow):
        for sheet2Row in range(1,sheet2MaxRow):
            cell1 = getCell(sheet1,sheet1Row,sheet1NumberColumn)
            cell2 = getCell(sheet2,sheet2Row,sheet2NumberColumn)
            if cell1.value == cell2.value:
                print(cell1.value)
        
    return
    
initInputExcel()
checkInputExcel()
procData()
#main()
