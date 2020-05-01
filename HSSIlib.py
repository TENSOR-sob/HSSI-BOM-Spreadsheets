_all_ = ['HSSI_SpliceLeftRightBySheetMark',
         'HSSI_SortedGirdLenByMark',
         'HSSI_MarkNetWt',
         'HSSI_ShipMarkByPointNumber',
         'HSSI_DictOfPointsBySheetNum',
         'HSSI_DictOfSheetNumByTenGirdMark',
         'HSSI_NsFsStiffMarkCountBySheetNumAndPoint',
         'HSSI_GetIntStiff',
         'HSSI_EndPointsBySheetNumber',
         'HSSI_FieldSplicePlMarkQtyByPoint',
         'HSSI_DictMatlSectByPoint']

import os

def HSSI_SpliceLeftRightBySheetMark(TensorJob):
    SpliceLeftRightBySheetMark = {}
    LeftMarkQtyDict = {}
    RightMarkQtyDict = {}
    # {'SheetNum1': {'GirdMark1': {'SplPlMk1': {'Left': Qty, 'Right': Qty}
    #                             {'SplPlMk2': {'Left': Qty, 'Right': Qty}
    #                             ..........
    #               {'GirdMark2': {'SplPlMk1': {'Left': Qty, 'Right': Qty}
    #                             {'SplPlMk2': {'Left': Qty, 'Right': Qty}
    #                             ..........
    # {'SheetNum2': {'GirdMark1': {'SplPlMk1': {'Left': Qty, 'Right': Qty}
    #                             {'SplPlMk2': {'Left': Qty, 'Right': Qty}
    #                             ..........
    #               {'GirdMark2': {'SplPlMk1': {'Left': Qty, 'Right': Qty}
    #                             {'SplPlMk2': {'Left': Qty, 'Right': Qty}
    #                             ..........

    FieldSplicePlMarksQtyByPointDict = HSSI_FieldSplicePlMarksQtyByPoint(TensorJob)
    if not bool(FieldSplicePlMarksQtyByPointDict):
        return SpliceLeftRightBySheetMark
    EndPointsBySheetNumberDict = HSSI_EndPointsBySheetNumber(TensorJob)
    DictOfSheetNumByTenGirdMark = HSSI_DictOfSheetNumByTenGirdMark(TensorJob)
    SortedGirdLenByMarkDict = HSSI_SortedGirdLenByMark(TensorJob)
    SpliceMatlSectByPointDict = HSSI_DictMatlSectByPoint(TensorJob)
    ShipMarkByPointNumberDict = HSSI_ShipMarkByPointNumber(TensorJob)
    UsedPntList = []
    FillVarName = [['webfp','LWBT','RWBT'], ['tffpl','LTFT','RTFT'], ['bffpl','LBFT','RBFT']]
    for GirdMark in SortedGirdLenByMarkDict:
        SheetNum = (DictOfSheetNumByTenGirdMark[GirdMark])
        Lpnt, Rpnt = EndPointsBySheetNumberDict[int(SheetNum)]
        SLpnt = str(Lpnt)
        SRpnt = str(Rpnt)
        LeftMarkQtyDict = {}
        RightMarkQtyDict = {}
        LRPntDict = {}
        if SLpnt not in UsedPntList:
            UsedPntList.append(SLpnt)
            LRPntDict['Left'] = SLpnt
        if SRpnt not in UsedPntList:
            UsedPntList.append(SRpnt)
            LRPntDict['Right'] = SRpnt
        SplPlMkDict = {}
        for LR in LRPntDict:
            Pnt = LRPntDict[LR]
            if Pnt in FieldSplicePlMarksQtyByPointDict:
                MarkQtyDict = FieldSplicePlMarksQtyByPointDict[Pnt]
                for SplPlMk in MarkQtyDict:
                    SplType = MarkQtyDict[SplPlMk][1]
                    Qty = MarkQtyDict[SplPlMk][0]
                    if SplType == 'FILL':
                        TempDict1 = SpliceMatlSectByPointDict[str(Pnt)]
                        TempList = []
                        for i in range(0,2):
                            if MarkQtyDict[SplPlMk][2] == FillVarName[i][0]:
                                LThk = float(TempDict1[FillVarName[i][1]])
                                RThk = float(TempDict1[FillVarName[i][2]])
                                Side = 'Left'
                                GirdSide = 'Right'
                                SideIndex = 0
                                if RThk < LThk:
                                    Side = 'Right'
                                    GirdSide = 'Left'
                                    SideIndex = 1
                                FillGirdMark = ShipMarkByPointNumberDict[int(Pnt)][SideIndex]
                                FillSheetNum = DictOfSheetNumByTenGirdMark[FillGirdMark]
                                if FillSheetNum not in SpliceLeftRightBySheetMark:
                                    SpliceLeftRightBySheetMark[FillSheetNum] = {FillGirdMark: {SplPlMk: {'Type': SplType, GirdSide: Qty}}}
                                else:
                                    if FillGirdMark not in SpliceLeftRightBySheetMark[FillSheetNum]:
                                        SpliceLeftRightBySheetMark[FillSheetNum][FillGirdMark] = {SplPlMk: {'Type': SplType, GirdSide: Qty}}
                                    else:
                                        if SplPlMk not in SpliceLeftRightBySheetMark[FillSheetNum][FillGirdMark]:
                                            SpliceLeftRightBySheetMark[FillSheetNum][FillGirdMark][SplPlMk] = {'Type': SplType, GirdSide: Qty}
                                        else:
                                            if GirdSide not in SpliceLeftRightBySheetMark[FillSheetNum][GirdMark][SplPlMk]:
                                                SpliceLeftRightBySheetMark[FillSheetNum][GirdMark][SplPlMk][GirdSide] = Qty
                                            else:
                                                PrevQty = SpliceLeftRightBySheetMark[FillSheetNum][GirdMark][SplPlMk][GirdSide]
                                                SpliceLeftRightBySheetMark[FillSheetNum][GirdMark][SplPlMk][GirdSide] = PrevQty + Qty

                    else:
                        if SheetNum not in SpliceLeftRightBySheetMark:
                            SpliceLeftRightBySheetMark[SheetNum] = {GirdMark: {SplPlMk: {'Type': SplType, LR: Qty}}}
                        else:
                            if GirdMark not in SpliceLeftRightBySheetMark[SheetNum]:
                                SpliceLeftRightBySheetMark[SheetNum][GirdMark] = {SplPlMk: {'Type': SplType, LR: Qty}}
                            else:
                                
                                if SplPlMk not in SpliceLeftRightBySheetMark[SheetNum][GirdMark]:
                                    SpliceLeftRightBySheetMark[SheetNum][GirdMark][SplPlMk] = {'Type': SplType, LR: Qty}
                                else:
                                    if LR not in SpliceLeftRightBySheetMark[SheetNum][GirdMark][SplPlMk]:
                                        SpliceLeftRightBySheetMark[SheetNum][GirdMark][SplPlMk][LR] = Qty
                                    else:
                                        PrevQty = SpliceLeftRightBySheetMark[SheetNum][GirdMark][SplPlMk][LR]
                                        SpliceLeftRightBySheetMark[SheetNum][GirdMark][SplPlMk][LR] = PrevQty + Qty
                    continue
#                    if GirdMark == 'G1D':
#                        print('A',SplPlMkDict)
#                    if SplPlMk in SplPlMkDict:
#                        if LR in SplPlMkDict[SplPlMk]:
#                            SplPlMkDict[SplPlMk][LR] = SplPlMkDict[SplPlMk][LR] + Qty
#                        else:
#                            SplPlMkDict[SplPlMk][LR] = Qty
#                    else:
#                        SplPlMkDict[SplPlMk] = {'Type': SplType, LR: Qty}
#                    if GirdMark == 'G1D':
#                        print(SplPlMkDict)

        
#        if not bool(SplPlMkDict):
#            continue
#        if SheetNum not in SpliceLeftRightBySheetMark:
#            SpliceLeftRightBySheetMark[SheetNum] = {GirdMark: SplPlMkDict}
#        else:
#            if GirdMark not in SpliceLeftRightBySheetMark[SheetNum]:
#                SpliceLeftRightBySheetMark[SheetNum][GirdMark] = SplPlMkDict
#            else:
#                if SplPlMk not in SpliceLeftRightBySheetMark[SheetNum][GirdMark]:
#                    SpliceLeftRightBySheetMark[SheetNum][GirdMark][SplPlMark] = {'Type': SplType, LR: 
#    print(SpliceLeftRightBySheetMark[906])

    return SpliceLeftRightBySheetMark
        
def HSSI_MarkNetWt(TensorJob):
    from os import listdir
    from os.path import isfile, join
    MarkNetWtDict = {}
    NetWtDir = 'J:/'+TensorJob+'/BILL/NETWTS'
    if not os.path.exists(NetWtDir):
        return MarkNetWtDict
    FileList = [f for f in listdir(NetWtDir) if isfile(join(NetWtDir, f))]
    for FN in FileList:
        with open(NetWtDir+'/'+FN, 'r') as f:
            NetWt = f.readline().strip('\n')
            MarkNetWtDict[FN] = NetWt
    return MarkNetWtDict

def HSSI_SortedGirdLenByMark(TensorJob):
    SortedGirdLenByMarkDict = {}
    LenDir = 'J:/'+TensorJob+'/BILL/SHIPWTS'
    for file in os.listdir(LenDir):
        if file.endswith(".len"):
            with open(LenDir+'/'+file,'r') as f:
                SortedGirdLenByMarkDict[file[:-4]] = float(f.readline().strip('\n'))
    return {k: v for k, v in sorted(SortedGirdLenByMarkDict.items(), key=lambda item: item[1])}

def HSSI_FieldSplicePlMarksQtyByPoint(TensorJob):
    FieldSplicePlMarksQtyByPointDict = {}
    FileName = 'splicemark'
    # List of splice keys, start col - 1, end col
    SRL = [['point', 0, 4],
           ['websp', 4, 10],
           ['webfp', 10, 16],
           ['tfospl', 16, 22],
           ['tfispl', 22, 28],
           ['tffpl', 28, 34],
           ['bfospl', 34, 40],
           ['bfispl', 40, 46],
           ['bffpl', 46, 52]
           ]
    FillList = ['webfp', 'tffpl', 'bffpl']
    DblQty = ['websp', 'webfp', 'tfispl', 'bfispl']
    try:
        splicemarkObj = open('J:/'+TensorJob+'/REF/'+FileName,'r')
    except:
        print(FileName + ' Not found in REF directory')
        return FieldSplicePlMarksQtyByPointDict
        
    i = 0
    for FileLine in splicemarkObj:
        SpliceDict = {}
        for Alist in SRL:
            SplID = Alist[0]
            SC = Alist[1]
            EC = Alist[2]
            if SplID == 'point':
                PointStr = FileLine[SC:EC].strip()
                continue
            Str = FileLine[SC:EC].strip()
            if len(Str) < 1:
                continue
            QTY = 1
            if SplID in DblQty:
                QTY = 2
            SKey = Str
            SplType = 'SPL'
            if SplID in FillList:
                SplType = 'FILL'
            if SKey in SpliceDict:
                SpliceDict[SKey][0] = SpliceDict[SKey][0] + QTY
            else:
                SpliceDict[SKey] = [QTY, SplType]
                if SplType == 'FILL':
                    SpliceDict[SKey] = [QTY, SplType, SplID]
            i = i + 1
        if PointStr in FieldSplicePlMarksQtyByPointDict:
            TempDict = FieldSplicePlMarksQtyByPointDict[Point]
            for SplMark in SpliceDict:
                AddQty = SpliceDict[SplMark][0]
                PrevQty = 0
                if SplMark in TempDict:
                    PrevQty = FieldSplicePlMarksQtyByPointDict[PointStr][SplMark][0]
                FieldSplicePlMarksQtyByPointDict[PointStr][SplMark][0] = PrevQty + AddQty
        else:
            FieldSplicePlMarksQtyByPointDict[PointStr] = SpliceDict
                
    return FieldSplicePlMarksQtyByPointDict

def HSSI_SortedGirdLenByMark(TensorJob):
    SortedGirdLenByMarkDict = {}
    LenDir = 'J:/'+TensorJob+'/BILL/SHIPWTS'
    for file in os.listdir(LenDir):
        if file.endswith(".len"):
            with open(LenDir+'/'+file,'r') as f:
                SortedGirdLenByMarkDict[file[:-4]] = float(f.readline().strip('\n'))
    return {k: v for k, v in sorted(SortedGirdLenByMarkDict.items(), key=lambda item: item[1])}

def HSSI_EndPointsBySheetNumber(TensorJob):
    SheetNumberEndPointsDict = {}
    SheetPointsDict = HSSI_DictOfPointsBySheetNum(TensorJob)
    SheetNumbers = SheetPointsDict.keys()
    for Sheet in SheetNumbers:
        SheetNumberEndPointsDict[Sheet] = [SheetPointsDict[Sheet][0],
                                           SheetPointsDict[Sheet][-1]]
    return SheetNumberEndPointsDict

def HSSI_ShipMarkByPointNumber(TensorJob):
    GirdByPointDict = {}
    GirdPtsObj = open('J:/'+TensorJob+'/REF/girderpts','r')
    for PointsLine in GirdPtsObj:
        GirdMark = PointsLine[1:8].strip()
        NumPts = int(PointsLine[10:12])
        StartPos = 13
        for i in range(1, NumPts+1):
            SP = StartPos+(i)*4
            EP = SP+3
            Point = abs(int(PointsLine[SP:EP]))
            if Point not in GirdByPointDict.keys():
                GirdByPointDict[Point] = [GirdMark]
            else:
                GirdByPointDict[Point].append(GirdMark)
    return GirdByPointDict

def HSSI_DictOfPointsBySheetNum(TensorJob):
    PointsList = []
    SheetPointsDict = {}
    SheetNumByTenGirdMarkDict = HSSI_DictOfSheetNumByTenGirdMark(TensorJob)
    GirdPtsObj = open('J:/'+TensorJob+'/REF/girderpts','r')
    for PointsLine in GirdPtsObj:
        GirdMark = PointsLine[1:8].strip()
        SheetNum = SheetNumByTenGirdMarkDict[GirdMark]
        NumPts = int(PointsLine[10:12])
        StartPos = 13
        for i in range(1, NumPts+1):
            SP = StartPos+(i)*4
            EP = SP+3
            Point = abs(int(PointsLine[SP:EP]))
            if SheetNum not in SheetPointsDict.keys():
                SheetPointsDict[SheetNum] = [Point]
            else:
                SheetPointsDict[SheetNum].append(Point)
    return SheetPointsDict
                            
def HSSI_DictOfSheetNumByTenGirdMark(TensorJob):
    SheetNumByTenGirdMarkDict = {}
    MarklookupObj = open('J:/'+TensorJob+'/REF/marklookup','r')
    for FileLine in MarklookupObj:
        Mark = FileLine[1:8].strip()
        if len(Mark) < 1:
            continue
        SheetNum = int(FileLine[21:28].strip())
        SheetNumByTenGirdMarkDict[Mark] = SheetNum
    return SheetNumByTenGirdMarkDict

def HSSI_DictOfFabMarkByTenMark(TensorJob):
    TenMarkByFabMarkDict = {}
    MarklookupObj = open('J:/'+TensorJob+'/REF/marklookup','r')
    for FileLine in MarklookupObj:
        Mark = FileLine[1:8].strip()
        if len(Mark) < 1:
            continue
        FabMark = FileLine[8:16].strip()
        TenMarkByFabMarkDict[Mark] = FabMark
    return TenMarkByFabMarkDict

def HSSI_GetIntStiff(TensorJob):
    SheetNumList = []
    MarkNsFsQtyList = []
    try:
        StfmarkObj = open('J:/'+TensorJob+'/REF/stfmark','r')
    except:
        return
    MarkPointDict = HSSI_ShipMarkByPointNumber(TensorJob)
    SheetMarkDict = HSSI_DictOfSheetNumByTenGirdMark(TensorJob)
    Point = 0
    MarkSideList = []
    for inline in StfmarkObj:
        Point = Point + 1
        try:
            Qty = int(inline[1:4])
        except:
            continue
        if Qty > 0:
            Jump = False
            try:
                TenGirdMark = MarkPointDict[Point][0]
            except:
                continue
            try:
                SheetNum = SheetMarkDict[TenGirdMark]
            except:
                continue
            for i in range(0, Qty):
                SM = i*12+5
                EM = SM+7
                SS = SM+9
                Mark = inline[SM:EM].strip()
                Side = inline[SS]
                if len(Mark) < 1 or len(Side) < 0:
                    Jump = True
                    continue
                MarkSideList.append([SheetNum,Mark,Side])
            if Jump:
                continue

            if SheetNum in SheetNumList:
                TempDict = {}
                Pos = SheetNumList.index(SheetNum)
                TempDict = list(MarkNsFsQtyList)[Pos]
    return(MarkSideList)

def HSSI_NsFsStiffMarkCountBySheetNumAndPoint(TensorJob,SheetNumByTenGirdMarkDict):

    MarkPointDict = HSSI_ShipMarkByPointNumber(TensorJob)
    SheetMarkDict = SheetNumByTenGirdMarkDict
    FabMarkByTenMarkDict = HSSI_DictOfFabMarkByTenMark(TensorJob)
    CfconnmarkObj = open('J:/'+TensorJob+'/REF/cfconnmark','r')
    MarkNsFsQtyDict = {}  #{Mark : [NS Qty, FS Qty]}
    SheetNumMarkNsFsQtyDict = {} #{ShtNum : [{Mark : [NS Qty, FS Qty]}, ... ]}
    SheetNumList = []
    MarkNsFsQtyList = []
    Point = 0
    for inline in CfconnmarkObj:
        Point = Point + 1
        MarkNsFsQtyDict = {}
        if len(inline[1:12].strip())>0 and Point != 1 and inline[1:12].strip() != '999':
            TenGirdMark = FabMarkByTenMarkDict[MarkPointDict[Point][0]]
            SheetNum = SheetMarkDict[str(TenGirdMark)]
            NsMark = inline[0:6].strip()
            FsMark = inline[6:12].strip()
            NlBrgMark = inline[76:82].strip()
            FlBrgMark = inline[82:88].strip()
            NrBrgMark = inline[88:94].strip()
            FrBrgMark = inline[94:100].strip()

            MarkList = [NsMark, FsMark, NlBrgMark, FlBrgMark, NrBrgMark, FrBrgMark]

            if SheetNum not in SheetNumMarkNsFsQtyDict:
                SheetNumMarkNsFsQtyDict[SheetNum] = {}
            NsFs = 1
            for PieceMark in MarkList:
                if NsFs == 1:
                    NsFs = 0
                else:
                    NsFs = 1
                if len(PieceMark.strip()) == 0:
                    continue
                PrevQtyList = [0, 0]
                if PieceMark not in SheetNumMarkNsFsQtyDict[SheetNum]:
                    PrevQtyList[NsFs] = 1
                else:
                    PrevQtyList = SheetNumMarkNsFsQtyDict[SheetNum][PieceMark]
                    PrevQtyList[NsFs] = PrevQtyList[NsFs] + 1
                    
                SheetNumMarkNsFsQtyDict[SheetNum][PieceMark] = PrevQtyList
    IntStiffData = HSSI_GetIntStiff(TensorJob)
    if len(IntStiffData) < 1:
        return(SheetNumMarkNsFsQtyDict)
    for IntStiffRec in IntStiffData:
        SheetNum = IntStiffRec[0]
        Mark = IntStiffRec[1]
        NsFs = IntStiffRec[2]
        if bool(SheetNumMarkNsFsQtyDict):
            SheetNumList = SheetNumMarkNsFsQtyDict.keys()
        NsQty = 0
        FsQty = 0
        if str(SheetNum) in SheetNumList:
            MarkNsFsQtyDict = SheetNumMarkNsFsQtyDict[str(SheetNum)]
            if Mark in MarkNsFsQtyDict:
                NsQty = MarkNsFsQtyDict[Mark][0]
                FsQty = MarkNsFsQtyDict[Mark][1]
        if NsFs == 'n':
            NsQty = NsQty + 1
        if NsFs == 'f':
            FsQty = FsQty + 1
        SheetNumMarkNsFsQtyDict[str(SheetNum)][Mark] = [NsQty,FsQty]
    return(SheetNumMarkNsFsQtyDict)

def HSSI_DictMatlSectByPoint(TensorJob):
    MatlSectByPointDict = {} #{'Point': {'LWBW': 'Dec Width'}
                             #          {'LWBT': 'Dec Thk'}
                             #          {'LTFW': 'Dec Width'}
                             #          {'LTFT': 'Dec Thk'}
                             #          {'LBFW': 'Dec Width'}
                             #          {'LBFT': 'Dec Thk'}
                             #          {'RWBW': 'Dec Width'}
                             #          {'RWBT': 'Dec Thk'}
                             #          {'RTFW': 'Dec Width'}
                             #          {'RTFT': 'Dec Thk'}
                             #          {'RBFW': 'Dec Width'}
                             #          {'RBFT': 'Dec Thk'}
                             # etc.
                             #}
    VS1 = ['L','R']
    VS2 = ['WB','TF','BF']
    VS3 = ['W','T']
    VS4 = ['In','Si']
    STC = [0,10,12,14,16,20,22,24,26,30,32,34,36,46,
           48,50,52,56,58,60,62,66,68,70,72]
    CT = 0
    SS = []
    for V1 in VS1:
        for V2 in VS2:
            for V3 in VS3:
                VAL = 0.0
                for V4 in VS4:
                    VN = V1 + V2 + V3
                    ST = STC[CT]
                    EN = STC[CT+1]
                    SS.append([VN,ST,EN])
                    CT = CT + 1
    MatlsectionObj = open('J:/'+TensorJob+'/REF/matlsection','r')

    PT = 0
    SP = [0,1]
    for inline in MatlsectionObj:
        PT = PT + 1
        IsZero = True
        TempDict = {}
        TempDict1 = {}
        for i in range(0,len(SS)-1,2):
            VN = SS[i][0]
            Str1 = inline[SS[i][1]:SS[i][2]].strip()
            if len(Str1) > 0:
                VAL1 = float(Str1)/12.0
            else:
                VAL1 = 0

            Str2 = inline[SS[i+1][1]:SS[i+1][2]].strip()
            if len(Str2) > 0:
                VAL2 = float(Str2)/192.0
            else:
                VAL2 = 0
#            VAL1 = float(inline[SS[i][1]:SS[i][2]])/12.0
#            VAL2 = float(inline[SS[i+1][1]:SS[i+1][2]])/192.0
            VAL = VAL1 + VAL2
            if VAL > 0.0001:
                IsZero = False
            TempDict1[VN] = format(VAL1 + VAL2, '.4f')
            TempDict[str(PT)] = TempDict1
        if not IsZero:    
            MatlSectByPointDict[str(PT)] = TempDict[str(PT)]

    return MatlSectByPointDict




            
    
        
         
         
        
    
