from openpyxl import Workbook

def addList(l,d):
    l.append(d)

# wb = getWb()
# wb.active
def getWb():
    return Workbook()

def activeWb(wb):
    return wb.active

def addWs(wb,name):
    return wb.create_sheet(name)

# add first addShtPos(wb,name,0):
def addWsPos(wb,name,pos):
    return wb.create_sheet(name,pos)

def getWsByName(wb,name):
    return wb[name]

def getWsName(ws):
    return ws.title

def getWsNames(wb):
    return wb.sheetnames

def copyWs(wb,ws):
    return wb.copy_worksheet(ws)

def getWsNamesLoop(wb):
    names = []
    for ws in wb:
        addList(names,getWsName(ws))
    return names

def setVal(ws,col,row,val):
    ws[col+row]=val

def setVal(ws,colrow,val):
    ws[colrow]=val
    
def getVal(ws,col,row):
    
    return ws[col+row]

def setWsName(ws,name):
    ws.title = name

# setShtTabColor(ws,"1072BA")
def setWsTabColor(ws,color):
    ws.sheet_properties.tabColor = color

def saveWb(wb,filePath):
    wb.save(filePath)
