import sys
sys.path.append("..")

from lib.utils import *

    

def main():
    wb = getWb()
    ws = activeWb(wb)
    setWsName(ws,"aaa")
    ws1 = addWs(wb,"bbb")

    setVal(ws1,"A3","dasdsa")

    copyWs(wb,ws1)
    saveWb(wb,"test.xlsx")
    
    


main()
