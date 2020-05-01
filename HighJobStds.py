#!/unixvol/opt/gcc/gcc-5.3.0/bin/python
import platform, os, sys, csv, pandas as pd, numpy as np, win32com.client, time
import HSSIlib
from os import path
try:
    platform = platform.system()
except:
    platform == 'Unix'
    
if platform == 'Windows':
    DriveStr = 'J:/'
else:
    DriveStr = '/users/jobs/'
    RootPath = os.path.abspath(os.path.normpath('/users/jobs/bin/python/openpyxl-2.5.0b1'))
    if RootPath not in sys.path:
        sys.path.append(RootPath)
#        import openpyxl


ModulesPath = os.path.abspath(os.path.normpath(DriveStr + 'bin/lib/HighJobStds'))
if ModulesPath not in sys.path:
    sys.path.append(ModulesPath)
# import GHlib, GHmmo, GHfit

TensorLibPath = os.path.abspath(os.path.normpath(DriveStr + 'bin/lib/TenLib'))
if TensorLibPath not in sys.path:
    sys.path.append(TensorLibPath)
import TFten

if platform == 'Windows':
    TenIn = input("Enter Job Number ")
else:
    print("Do not Enter Job Number")
    TenIn = sys.stdin.readline()
    
try:
    TensorJob = str(TenIn).strip()
except:
    TensorJob = TenIn
PathName = DriveStr + TensorJob + '/BILL/PRODUCT'
print(PathName)

try:
    os.chdir(PathName)
except:
    print('Job ' + TensorJob + ' does not exist')
    sys.exit()


ProductDataFileName = 'product.csv'

try:
    ProductDataObject = open(ProductDataFileName, 'r')
except:
    print(PathName + '/' + ProductDataFileName + ' does not exist')
    sys.ext()

def openWorkbook(xlapp, xlfile):
    try:
        xlwb = xlapp.Workbooks(xlfile)
    except Exception as e:
        print(e)
        xlwb = None
    return(xlwb)

JobStdWbFileName2 = 'JobStandards-tmp.xlsm'
JobStdWbFileName = 'JobStandards.xlsm' 
if path.exists(PathName + '/' + JobStdWbFileName2):
    if path.exists(PathName + '/' + JobStdWbFileName):
        os.remove(PathName + '/' + JobStdWbFileName)
    os.rename(PathName + '/' + JobStdWbFileName2,PathName + '/' + JobStdWbFileName)
    
JobStdWbSheetNameList = []
try:
    excel = win32com.client.Dispatch('Excel.Application')

    JobStdWb = excel.Workbooks.Open(PathName + '/' + JobStdWbFileName)

    excel.Visible = True
    JobStdWb.Unprotect()
    DelSheetList = ("X1", "X2", "X3", "X4", "M1", "Z1", "Special")
    SheetsToKeep = ("Template", "Total", "L Weights", "Special")
    SheetList = []
    for ws in JobStdWb.Sheets:
        SheetList.append(ws.Name)
    excel.DisplayAlerts = False
    for ShName in SheetList:
        if ShName not in SheetsToKeep:
            JobStdWb.Sheets(ShName).Delete()
    excel.DisplayAlerts = True
    for sh in JobStdWb.Sheets:
        if sh.Name not in JobStdWbSheetNameList:
            JobStdWbSheetNameList.append(sh.Name)
    excel.Sheets("Special").Select()
    excel.Range("C4:U43").Select()
    try:
        excel.Selection.SpecialCells(2).ClearContents()
    except:
        print("")
        
    from csv import DictReader

    MarkNetWtDict = HSSIlib.HSSI_MarkNetWt(TensorJob)

    SpecialList = ['Nut', 'HS Bolt', 'Std Wash', 'Stud']
    SpecialDict = {}
    with ProductDataObject as read_obj:
        dict_reader = DictReader(read_obj)
        list_of_dict = list(dict_reader)
        a = 0
        HeaderLocDict = {}
        HeadersDict = {}
        count = 0
        for item in list_of_dict:
            if str(item['DWG'][0]).isalpha():
                if item['DWG'] not in JobStdWbSheetNameList:
                    JobStdWbSheetNameList.append(item['DWG'])
                    WBT = JobStdWb.Worksheets("Template")
                    excel.DisplayAlerts=False
                    WBT.Copy(Before=WBT)
                    excel.DisplayAlerts=True
                    WBN = JobStdWb.ActiveSheet
                    WBN.Name = item['DWG']
                    WBN.Unprotect()
                    PleaseStop = False
                    for col in range(1, 22):
                        if WBN.Cells(2,col).Value is None:
                            continue
                        if WBN.Cells(2,col).Value == "WT. EA.":
                            HeaderLocDict["WT. EA."] = col
                            continue
                        if str(WBN.Cells(2,col).Value) == "COMMODITY":
                            HeaderLocDict["COMMODITY"] = col
                            continue
                        if str(WBN.Cells(2,col).Value) == "LENGTH":
                            HeaderLocDict["LENGTH"] = col
                            continue
                        if str(WBN.Cells(2,col).Value) == "DEDUCT WT FROM EACH":
                            HeaderLocDict["DEDUCT WT FROM EACH"] = col
                            continue
                        for Header in item.keys():
                            if WBN.Cells(2,col).Value == Header:
                                HeaderLocDict[Header] = col
                                PleaseStop = True
                                break
                    HeadersDict[item['DWG']] = HeaderLocDict
            else:
                if item['COMM'] in SpecialList:
                    count = count + 1
                    if item['MARK'] in SpecialDict.keys():
                        SpecialDict[item['MARK']]['QTY'] = str(int(SpecialDict[item['MARK']]['QTY']) + int(item['QTY']))
                        SpecialDict[item['MARK']]['DWG'] = 'Special'
                    else:
                        SpecialDict[item['MARK']] = item
        TempDict = {}
        count = 0
        for Mark in sorted(set(SpecialDict.keys())):
#            SpecialDict = item[Mark]
            SpecialDict[Mark]["DWG"] = "Special"
            HeadersDict["Special"] = HeaderLocDict
#            list_of_dict.append(SpecialDict)

        CommList1 = ("WTM", "WT", "MT", "ST", "MC", "HP", "HPT", "W", "M", "S", "C")
        CommList2 = ("L", "HSS", "Tube")
        CommList3 = ("BR", "CP", "FL", "PL", "SQ", "PAD", "SL", "Elast Brg", "Fab Pad")
        CommList4 = ("CPG", "GA", "SHT")
        CommList5 = ("HSB", "TCB", "LJB", "BLT", "MB", "CS", "SD", "RD", "RDE", "SDE", "HS Bolt", "RB", "Stud", "Anch Bolt", "Screw")
        CommList6 = ("NU", "HN", "HHN", "LN", "JN", "WA", "HSW", "BW", "CW", "LIW", "LW", "Nut", "Std Wash")
        CommList7 = ("SW", "Anch Wash")
        CommList8 = ("P", "PX", "PXX", "PS", "Pipe")
        CommList9 = ("RB", "Rebar")
        CommList5_6_8 = CommList5 + CommList6 + CommList8
        HeadersLocDic = {}

        for BillName in HeadersDict.keys():
            HeaderLocDict = HeadersDict[BillName]
            Count = 0
            Row = 3
            SpecialMarkUsedList = []
            for MatlItem in list_of_dict:
                if MatlItem['DWG'] == BillName:
                    if BillName == 'Special' and MatlItem['MARK'] in SpecialMarkUsedList:
                        continue
                    else:
                        SpecialMarkUsedList.append(MatlItem['MARK'])
                    JobStdWb.Sheets(BillName).Activate()
                    Row = Row + 1
                    if Row > 41:
                        RangeStrCopy = str("A"+str(Row)+":"+"AC"+str(Row))
                        RangeStrPaste = str("A"+str(Row+1)+":"+"AC"+str(Row+1))
                        rangeObj = JobStdWb.Sheets(BillName).Range(RangeStrCopy)
                        rangeObj.EntireRow.Insert()
                        RangeStrCopy = str("A"+str(Row+1)+":"+"AC"+str(Row+1))
                        RangeStrPaste = str("A"+str(Row)+":"+"AC"+str(Row))
                        CopyObj = JobStdWb.Sheets(BillName).Range(RangeStrCopy)
                        PasteObj = JobStdWb.Sheets(BillName).Range(RangeStrPaste)
                        CopyObj.Copy(PasteObj)
                    Comm = ''
                    for ColHeader in HeaderLocDict.keys():

                        if str(ColHeader) == 'COMMODITY':
                            JobStdWb.Sheets(BillName).Cells(Row,HeaderLocDict['COMMODITY']).Value = MatlItem['COMM']
                        elif str(ColHeader) == 'WT. EA.':
                            JobStdWb.Sheets(BillName).Cells(Row,HeaderLocDict['WT. EA.']).Value = MatlItem['WGT']
                        elif str(ColHeader) == 'DESCRIPTION' and str(MatlItem['COMM']) in CommList1:
                            Xpos = str(MatlItem['DESCRIPTION']).find("x")
                            str1 = (MatlItem['DESCRIPTION'][:Xpos])
                            val2 =(MatlItem['DESCRIPTION'][Xpos+1:])
                            JobStdWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]).Value = str1.strip()
                            JobStdWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]+2).Value = val2
                        elif str(ColHeader) == 'DESCRIPTION' and str(MatlItem['COMM']) in CommList2:
                            Xpos = str(MatlItem['DESCRIPTION']).find("x")
                            Str1 = MatlItem['DESCRIPTION'][:Xpos]
                            StrRem = MatlItem['DESCRIPTION'][Xpos+1:]
                            Xpos = StrRem.find("x")
                            Str2 = StrRem[:Xpos]
                            Str3 = StrRem[Xpos+1:]
                            if Str1.find("-") > 0:
                                Str1 = Str1.replace("-", " ")
                            if Str2.find("-") > 0:
                                Str2 = Str2.replace("-", " ")
                            if len(Str2) < 1 and (str(MatlItem['COMM']) == "TS" or str(MatlItem['COMM']) == "HSS" or str(MatlItem['COMM']) == "Tube"):
                                Str2 = Str1
                            if Str3.find("-") > 0:
                                Str3 = Str3.replace("-", " ")
                            JobStdWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]).Value = TFten.ftd(Str1)*12
                            JobStdWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]+2).Value = TFten.ftd(Str2)*12
                            JobStdWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]+4).Value = TFten.ftd(Str3)*12
                        elif str(ColHeader) == 'DESCRIPTION' and str(MatlItem['COMM']) in CommList3:
                            Xpos = str(MatlItem['DESCRIPTION']).find("x")
                            Str1 = MatlItem['DESCRIPTION'][:Xpos]
                            if Str1.find("-") > 0:
                                Str1 = Str1.replace("-", " ")
                            val1 = TFten.ftd(Str1)*12.0
                            Str2 = MatlItem['DESCRIPTION'][Xpos+1:]
                            if Str2.find("-") > 0:
                                Str2 = Str2.replace("-", " ")

                            val2 = TFten.ftd(Str2)*12.0
                            JobStdWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]).Value = val1
                            JobStdWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]+2).Value = val2
                        
                        elif str(ColHeader) == 'DESCRIPTION' and (str(MatlItem['COMM']) in CommList5_6_8):
                            Str1 = MatlItem['DESCRIPTION']
                            if Str1.find("-") > 0:
                                Str1 = Str1.replace("-", " ")
                            if Str1.find(" ROD") > 0:
                                Str1 = Str1.replace(" ROD", "")
                            val1 = TFten.ftd(Str1)*12.0
                            JobStdWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]).Value = val1
                        elif str(ColHeader) == 'LENGTH':
                            JobStdWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]).Value = MatlItem['LEN-FT']
                            JobStdWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]+1).Value = MatlItem['LEN-IN']
                        elif str(ColHeader) == 'DEDUCT WT FROM EACH':
                            if MatlItem['MARK'] in MarkNetWtDict:
                                Deduct = 0.0
                                if MatlItem['COMM'] == 'PL':
                                    Deduct = float(MatlItem['WGT']) - float(MarkNetWtDict[MatlItem['MARK']])
                                    if Deduct > 0.01:
                                        if float(MatlItem['WGT'])/Deduct < .01 or Deduct < 10.0:
                                            Deduct = 0.0
                                JobStdWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]).Value = str(Deduct)
                            else:
                                continue
                        else:
                            if BillName == "Special" and ColHeader != "QTY":
                                JobStdWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]).Value = str(MatlItem[ColHeader])
                            elif BillName != "Special":
                                JobStdWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]).Value = str(MatlItem[ColHeader])
        JobStdWb.Application.Run("BuildTotalSheet")
        JobStdWb.Save()
        JobStdWb.Close()
    excel.Quit()    
        


except Exception as e:
    print(e)
    
finally:
    # RELEASE RESOURCES
    JobStdWb = None
    excel = None
    
print('done')
