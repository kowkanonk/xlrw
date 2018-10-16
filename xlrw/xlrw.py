import xlrd
import xlutils.copy

inBook = xlrd.open_workbook('C:\\Users\\Administrator\\Desktop\\Reconciliation.xls', formatting_info=True)
outBook = xlutils.copy.copy(inBook)

def _getOutCell(outSheet, colIndex, rowIndex):
    """ HACK: Extract the internal xlwt cell representation. """
    row = outSheet._Worksheet__rows.get(rowIndex)
    if not row: return None

    cell = row._Row__cells.get(colIndex)
    return cell

def setOutCell(outSheet, col, row, value):
    """ Change cell value without changing formatting. """
    # HACK to retain cell style.
    previousCell = _getOutCell(outSheet, col, row)
    # END HACK, PART I

    outSheet.write(row, col, value)

    # HACK, PART II
    if previousCell:
        newCell = _getOutCell(outSheet, col, row)
        if newCell:
            newCell.xf_idx = previousCell.xf_idx
            print(newCell.xf_idx)
    # END HACK

outSheet = outBook.get_sheet(1)
setOutCell(outSheet, 0, 0, 'Test')
outBook.save('C:\\Users\\Administrator\\Desktop\\output.xls')
