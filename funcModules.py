def getTargetCell(strTargetCol, listCols):
    FirstCellInTargetCol = list(filter(lambda x: x.value == strTargetCol, listCols))[0]
    return FirstCellInTargetCol