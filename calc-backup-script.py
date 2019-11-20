import uno
from com.sun.star.sheet.CellFlags import (
    VALUE as NUM_VAL, DATETIME, STRING, FORMULA)
from collections import Counter

def copy_all_used_cells():
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet1 = oDoc.Sheets['All']
    CELLFLAGS = STRING | NUM_VAL | DATETIME | FORMULA  # any value
    oSheet1.clearContents(CELLFLAGS)
    MAX_COLS = 6  # up to column Z
    all_data = [[] for dummy in range(MAX_COLS)]
    for oSheet in oDoc.getSheets()[1:]:
        oRange = oSheet.getCellRangeByName("A2:C10000")
        oSheetCellRanges = oRange.queryContentCells(CELLFLAGS)
        xEnum = oSheetCellRanges.getCells().createEnumeration()
        while xEnum.hasMoreElements():
            oCell = xEnum.nextElement()
            oAddr = oCell.getCellAddress()
            all_data[oAddr.Column].append(oCell.getString())
    for col, col_data in enumerate(all_data):
        if col_data:
            col_data.sort()
            col_letter = chr(ord('A') + col)
            rangeName = "%s1:%s%d" % (
                col_letter, col_letter, len(col_data))
            oRange = oSheet1.getCellRangeByName(rangeName)
            data_tuples = ((val,) for val in col_data)  # 1 tuple per row
            oRange.setDataArray(data_tuples)

    BOLD = 150
    sheet = oDoc.Sheets['All']
    target = oDoc.Sheets['BackupStatus']
    cursor = target.createCursorByRange(target[0,0])
    cursor.collapseToCurrentRegion()
    cursor.clearContents(1023)

    cursor = sheet.createCursorByRange(sheet[0,0])
    cursor.collapseToCurrentRegion()
    data = cursor.DataArray

    for col in range(len(data[0])):
        values = tuple(zip(*data))[col]
        counter = Counter(filter(None, values))
        row = 0
        for k, v in counter.items():
            cell = target[row, col]
            cell.String = k
            if not v == 1:
                cell.CharWeight = BOLD
            row += 1
    return


g_exportedScripts = copy_all_used_cells,
