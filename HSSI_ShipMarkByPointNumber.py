def HSSI_ShipMarkByPointNumber(RetPoint,TensorJob):
    import os
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
    return GirdByPointDict[RetPoint]

def HSSI_DictOfPointsBySheetNum(TensorJob):
    import os
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
            if len(SheetPointsDict[SheetNum]) < 1:
                SheetPointsDict[SheetNum] = [Point]
            else:
                SheetPointsDict[SheetNum].append(Point)
    return SheetsPointsDict
                            
def HSSI_DictOfSheetNumByTenGirdMark(TensorJob):
    import os
    SheetNumByTenGirdMarkDict = {}
    MarklookupObj = open('J:/'+TensorJob+'REF/marklookup','r')
    for FileLine in MarklookupObj:
        Mark = FileLine[1:8].strip()
        SheetNum = int(FileLine[21:28].strip())
        SheetNumByTenGirdMarkDict[Mark] = SheetNum

def HSSI_NsFsCountBySheetAndMark(Sheet,Mark,NsFs,TensorJob):
    import os
    
        
         
         
        
    
