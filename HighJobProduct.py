3#!/unixvol/opt/gcc/gcc-5.3.0/bin/python
import platform, os, sys, csv, pandas as pd, numpy as np, win32com.client, re
import HSSIlib
from os import path
from collections import defaultdict
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

ModulesPath = os.path.abspath(os.path.normpath(DriveStr + 'bin/lib/HighJobStds'))
if ModulesPath not in sys.path:
    sys.path.append(ModulesPath)

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
    list_of_dict = []
    with open(ProductDataFileName,'r') as infile:
        reader = csv.DictReader(infile)
        list_of_dict.append(list(reader))
except:
    print(PathName + '/' + ProductDataFileName + ' does not exist')
    sys.exit()
finally:
    infile.close()

SheetNumByTenGirdMarkDict = {}
Temp = list_of_dict[0]
for item in Temp:
   if len(str(item['COMM'])) == 0 and str(item['DWG'][0]).isdigit() and len(str(item['MARK'])) > 0:
        Mark = str(item['MARK']).strip()
        SheetNum = str(item['DWG']).strip()
        SheetNum = re.sub("[^0-9]", "",SheetNum)
        SheetNumByTenGirdMarkDict[Mark] = SheetNum
                                           
ProductWbFileName = 'Products.xlsx'
ProductWbFileName2 = 'Products-tmp.xlsx'
print(PathName + '/' + ProductWbFileName2)
if path.exists(PathName + '/' + ProductWbFileName2):
    if path.exists(PathName + '/' + ProductWbFileName):
        os.remove(PathName + '/' + ProductWbFileName)
    os.rename(PathName + '/' + ProductWbFileName2, PathName + '/' + ProductWbFileName)
MarkNetWtDict = HSSIlib.HSSI_MarkNetWt(TensorJob)
    
ProductWbSheetNameList = []
ProductGirderDict = {}
NsFsQtyDict = {}

try:
    NsFsQtyDict = HSSIlib.HSSI_NsFsStiffMarkCountBySheetNumAndPoint(TensorJob,SheetNumByTenGirdMarkDict)
except Exception as e:
    print(e)
    print("Failed Call to HSSI_NsFsStiffMarkCountBySheetNumAndPoint")
print("STARTING EXCEL")
try:
    excel = win32com.client.Dispatch('Excel.Application')
    print("OPENING WORKBOOK " + PathName + '/' + ProductWbFileName)
    ProductWb = excel.Workbooks.Open(PathName + '/' + ProductWbFileName)

    excel.Visible = True
    ProductWb.Unprotect()
    DelSheetList = ("X1", "X2", "X3", "X4", "M1", "Z1", "Special")
    SheetsToKeep = ("Template")
    SheetList = []
    for ws in ProductWb.Sheets:
        SheetList.append(ws.Name)
    excel.DisplayAlerts = False
    for ShName in SheetList:
        if ShName not in SheetsToKeep:
            ProductWb.Sheets(ShName).Delete()
    excel.DisplayAlerts = True
    for sh in ProductWb.Sheets:
        if sh.Name not in ProductWbSheetNameList:
            ProductWbSheetNameList.append(sh.Name)
    from csv import DictReader
    SpliceLeftRightByShipMarkDict = HSSIlib.HSSI_SpliceLeftRightBySheetMark(TensorJob)
    SpliceLRShipMarkDict = {}
#    for item in list_of_dict:
#        print(item)
#        exit()
#    with ProductDataObject as read_obj:
#        dict_reader = DictReader(read_obj)
#        list_of_dict = list(dict_reader)
    
    a = 0
    HeaderLocDict = {}
    HeadersDict = {}
    Temp = list_of_dict[0]
    for item in Temp:
        if str(item['DWG'][0]).isdigit():
            item['DWG'] = re.sub("[^0-9]","",item['DWG'])
            if item['PROD'] not in ProductGirderDict.keys():
                ProductGirderDict[item['PROD']] = [item['DWG']]
            else:
                if item['DWG'] not in ProductGirderDict[item['PROD']]:
                    ProductGirderDict[item['PROD']].append(item['DWG'])
            if item['DWG'] not in ProductWbSheetNameList:
                ProductWbSheetNameList.append(item['DWG'])
                WBT = ProductWb.Worksheets("Template")
                WBT.Copy(Before=WBT)
                WBN = ProductWb.ActiveSheet
                Str1 = re.sub("[^0-9]","",item['DWG'])
                WBN.Name = Str1
                WBN.Unprotect()
                PleaseStop = False
                for col in range(1, 30):
                    if WBN.Cells(2,col).Value is None:
                        continue
                    if str(WBN.Cells(2,col).Value) == "QTY EA.":
                        HeaderLocDict["QTY"] = col
                        continue
                    if str(WBN.Cells(2,col).Value) == "LENGTH":
                        HeaderLocDict["LENGTH"] = col
                        continue
                    if str(WBN.Cells(2,col).Value) == "NS":
                        HeaderLocDict["NS"] = col
                        continue
                    if str(WBN.Cells(2,col).Value) == "FS":
                        HeaderLocDict["FS"] = col
                        continue            
                    if str(WBN.Cells(2,col).Value) == "DEDUCT WT FROM EACH":
                        HeaderLocDict["DEDUCT WT FROM EACH"] = col
                        continue
                    if str(WBN.Cells(2,col).Value) == "Left End":
                        HeaderLocDict["Left End"] = col
                        continue
                    if str(WBN.Cells(2,col).Value) == "Right End":
                        HeaderLocDict["Right End"] = col
                        continue
                    for Header in item.keys():
                        if WBN.Cells(2,col).Value == Header:
                            HeaderLocDict[Header] = col
                            PleaseStop = True
                            break

                    HeadersDict[item['DWG']] = HeaderLocDict
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
        NsFsQtyMarkDict = {}
        excel.DisplayAlerts = False
        ShipRangeCount = 0
        ShipRangeDict = {}
        PrevBillName = ''
        TempRow = []
    for BillName in HeadersDict:
            HeaderLocDict = HeadersDict[BillName]
            Count = 0
            Row = 3
            ShipRangeDict[BillName] = {'SRow': [], 'ERow': 0}
            ShipMark = ''

            for MatlItem in Temp:

                TabDwgName = re.sub("[^0-9]","",MatlItem['DWG'])

                try:
                    NsFsQtyMarkDict = NsFsQtyDict[TabDwgName]
                except:
                    NsFsQtyMarkDict = {}
                if MatlItem['DWG'] == BillName:

                    Row = Row + 1

                    ShipRangeDict[BillName]['ERow'] = Row
                    if PrevBillName != BillName:
                        ProductWb.Sheets(BillName).Activate()
                        PrevBillName = BillName
                    if Row > 41:
                        RangeStrCopy = str("A"+str(Row)+":"+"AE"+str(Row))
                        RangeStrPaste = str("A"+str(Row+1)+":"+"AE"+str(Row+1))
                        rangeObj = ProductWb.Sheets(BillName).Range(RangeStrCopy)
                        rangeObj.EntireRow.Insert()
                        RangeStrCopy = str("A"+str(Row+1)+":"+"AE"+str(Row+1))
                        RangeStrPaste = str("A"+str(Row)+":"+"AE"+str(Row))
                        CopyObj = ProductWb.Sheets(BillName).Range(RangeStrCopy)
                        PasteObj = ProductWb.Sheets(BillName).Range(RangeStrPaste)
                        CopyObj.Copy(PasteObj)
                    if len(MatlItem['MARK']) > 0 and len(MatlItem['COMM']) == 0:
                        ShipMark = MatlItem['MARK']
                        TempRow = ShipRangeDict[BillName]['SRow']
                        TempRow.append(Row)
                        ShipRangeDict[BillName]['SRow'] = TempRow
                        if Row > 4:
                            MergeRange = str("E"+str(Row)+":"+"G"+str(Row))
                            ProductWb.Sheets(BillName).Range(MergeRange).Merge()
                    SheetNum = int(BillName)
                    if SheetNum in SpliceLeftRightByShipMarkDict:
                        if ShipMark in SpliceLeftRightByShipMarkDict[SheetNum]:
                            SpliceLRShipMarkDict = SpliceLeftRightByShipMarkDict[SheetNum][ShipMark]


                    Comm = ''
                    for ColHeader in HeaderLocDict.keys():
                        if str(ColHeader) == 'COMM':
                            ProductWb.Sheets(BillName).Cells(Row,HeaderLocDict['COMM']).Value = MatlItem['COMM']
                        elif str(ColHeader) == 'DESCRIPTION' and str(MatlItem['COMM']) in CommList1:
                            Xpos = str(MatlItem['DESCRIPTION']).find("x")
                            str1 = (MatlItem['DESCRIPTION'][:Xpos])
                            val2 = float(MatlItem['DESCRIPTION'][Xpos+1:])
                            ProductWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]).Value = str1.strip()
                            ProductWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]+2).Value = val2
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
                            ProductWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]).Value = TFten.ftd(Str1)*12
                            ProductWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]+2).Value = TFten.ftd(Str2)*12
                            ProductWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]+4).Value = TFten.ftd(Str3)*12
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
                            ProductWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]).Value = val1
                            ProductWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]+2).Value = val2
                          
                        elif str(ColHeader) == 'DESCRIPTION' and (str(MatlItem['COMM']) in CommList5_6_8):
                            Str1 = MatlItem['DESCRIPTION']
                            if Str1.find("-") > 0:
                                Str1 = Str1.replace("-", " ")
                            if Str1.find(" ROD") > 0:
                                Str1 = Str1.replace(" ROD", "")
                            val1 = TFten.ftd(Str1)*12.0
                            ProductWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]).Value = val1
                        elif str(ColHeader) == 'LENGTH':
                            ProductWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]).Value = MatlItem['LEN-FT']
                            ProductWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]+1).Value = MatlItem['LEN-IN']
                        elif str(ColHeader) == 'NS':

                            if str(MatlItem['MARK']) in NsFsQtyMarkDict.keys():
                                ProductWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]).Value = NsFsQtyMarkDict[str(MatlItem['MARK'])][0]
                        elif str(ColHeader) == 'FS':
                            if str(MatlItem['MARK']) in NsFsQtyMarkDict.keys():
                                ProductWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]).Value = NsFsQtyMarkDict[str(MatlItem['MARK'])][1]
                        elif str(ColHeader) == 'DEDUCT WT FROM EACH':
                            NMark = ShipPieceMark = ShipMark+MatlItem['MARK']
                            T1 = MatlItem['MARK'] in MarkNetWtDict
                            T2 = ShipPieceMark in MarkNetWtDict
                            if T1:
                                NMark = MatlItem['MARK']
                            if T1 or T2:
                                Deduct = 0.0
                                if MatlItem['COMM'] == 'PL':
                                    Deduct = float(MatlItem['WGT']) - float(MarkNetWtDict[NMark])
                                    if Deduct > .01:
                                        if float(MatlItem['WGT'])/Deduct < 0.01 or Deduct < 10.0:
                                            Deduct = 0.0
                                ProductWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]).Value = str(Deduct)
#                            else:
#                                continue
     
                        elif str(ColHeader) == 'Left End':
                            if str(MatlItem['MARK']) in SpliceLRShipMarkDict:
                                if 'Left' in SpliceLRShipMarkDict[str(MatlItem['MARK'])]:
                                    Qty = SpliceLRShipMarkDict[str(MatlItem['MARK'])]['Left']
                                    ProductWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]).Value = Qty
                        elif str(ColHeader) == 'Right End':
                            if str(MatlItem['MARK']) in SpliceLRShipMarkDict:
                                if 'Right' in SpliceLRShipMarkDict[str(MatlItem['MARK'])]:
                                    Qty = SpliceLRShipMarkDict[str(MatlItem['MARK'])]['Right']
                                    ProductWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]).Value = Qty
                        else:
                            ProductWb.Sheets(BillName).Cells(Row,HeaderLocDict[ColHeader]).Value = MatlItem[ColHeader]
    for BillName in ShipRangeDict:
            NF = len(ShipRangeDict[BillName]['SRow'])
            ERow = ShipRangeDict[BillName]['ERow']
            Count = 0
            FC = ProductWb.Sheets(BillName).Range("A2:AA2").Find("SHIP WT. EA. (LBS)")
            for SR in ShipRangeDict[BillName]['SRow']:
                Count = Count + 1
                ER = ERow
                if Count < NF:
                    ER = ShipRangeDict[BillName]['SRow'][Count] - 1
                FormulaStr = '=IF(B'+str(SR)+'="","",SUM(V'+str((SR+1))+':V'+str((ER))+'))'
                ProductWb.Sheets(BillName).Cells(SR,FC.Column).Formula = FormulaStr
    excel.DisplayAlerts = True
    print("SAVING WORKBOOKS")

    ProdList = ProductGirderDict.keys()
    WbObjDict = {}
    if len(ProdList) > 0:
            for Prod in ProdList:
                print("CLOSING OLD WORKBOOKS")
                if os.path.isfile(PathName + '/' + Prod + '.xlsx'):
                    os.remove(PathName + '/' + Prod + '.xlsx')
                ProductWb.SaveAs(PathName + '/' + Prod + '.xlsx')
            print("CLOSING PRODUCT WORKBOOK ", Prod, '.xlsx')
            ProductWb.Close(True)
            for Prod in ProdList:
                print("OPENING WORKBOOKS")
                try:
                    excel = win32com.client.GetActiveObject('Excel.Application')
                    print("Running Excel Found")
                    excel
                except:
                     excel = new_Excel(visible=visible)
                wb = excel.Workbooks.Open(PathName + '/' + Prod + '.xlsx')
                excel.DisplayAlerts = False
                print("DELETING EXTRA WORKSHEETS")
                ShNameList = []
                for ws in wb.Sheets:
                    ShNameList.append(ws.Name)
                for ShName in ShNameList:
                    if str(ShName) not in ProductGirderDict[Prod] and str(ShName) != 'Template':
                        wb.Sheets(ShName).Delete()
                        print("Deleting ", ShName, " in ", wb.FullName)
                excel.DisplayAlerts = True
                wb.RefreshAll()
                wb.Close(True)
    excel.Quit()
        


except Exception as e:
    print(e)
    
finally:
    # RELEASE RESOURCES
    JobStdWb = None
    excel = None
    
print('done')
