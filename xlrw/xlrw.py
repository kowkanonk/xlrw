###############################################################################
#
# Format - A class for writing the Excel XLSX Worksheet file.
#
# Copyright 2018, Weixl, cd8866@qq.com
#

import xlrd             # http://pypi.python.org/pypi/xlrd
import xlutils.copy     # http://pypi.python.org/pypi/xlutils
                        # http://pypi.python.org/pypi/xlwt
import sys

class xlrw:
    """
    A class for writing the Excel XLSX Format file using existed Format.

    """

    ###########################################################################
    #
    # Public API.
    #
    ###########################################################################

    def __init__(self, tmplBook, filename):
        """
        Constructor. 

        Args:
            tmplBook: file path of template excel file.
            filename: file path of result excel file

        Returns:
            Nothing.

        """
        self.__tmplBook = xlrd.open_workbook(tmplBook, formatting_info=True)
        self.__filename = filename

        # get a copy of excel book used for data manipulation
        self.__resultBook = xlutils.copy.copy(self.__tmplBook)


    def set_out_cell(self, outSheet, row, col, value):
        """
        set cell value but retain cell stype for specific sheet

        Args:
            outSheet: sheet of workbook
            col: The cell column (zero indexed).
            row: The cell row (zero indexed).
            value: new cell value

        Returns:
            Nothing.

        """

        # HACK to retain cell style.
        previousCell = self._getOutCell(outSheet, row, col)
        # END HACK, PART I

        outSheet.write(row, col, value)

        # HACK to apply retained cell style for new cell
        if previousCell:
            newCell = self._getOutCell(outSheet, row, col)
            if newCell:
                newCell.xf_idx = previousCell.xf_idx
        # END HACK PART II

        self.__resultBook.save(self.__filename)

    def set_out_cells(self, outSheet, dict):
        """
        Set up sheet cells via dict.

        Args:
            dict: { {row: col}: value, {1,1}: 'cell_value'}

        Returns:
            Nothing.

        """

    def get_workbook(self):
        """
        get result workbook for sheets modification.
        primary method of xlrd workbook:
            ! sheet_by_index(sheet_index): return a sheet
            ! sheet_names(): return list of sheet names
            ! sheet_by_name(sheet_name): return a sheet
            get_sheet(sheet_index): return a sheet

        Args:
            Nothing.

        Returns:
            result workbook

        """
        return self.__resultBook

    def sheet_by_index(self, sheet_index):
        return self.__resultBook.get_sheet(sheet_index)


    def sheet_by_name(self, sheet_name):
        index = self.__resultBook.sheet_index(sheet_name)
        return self.__resultBook.get_sheet(index)

    def load_cell(self, filename):
        """
        get cell value from file, then loaded into dict

        Args:
            filename: config/data file which cell values are stored.

        Returns:
            result: dict with cell value, { {row: col}: value, {1,1}: 'cell_value'}

        """
        pass

    def set_cell_format(self, row, col, format):
        """
        stub

        """
        pass

    ###########################################################################
    #
    # Private API.
    #
    ###########################################################################

    def _getOutCell(self, outSheet, rowIndex, colIndex):
        """
        get cell in the outSheet of workbook

        Args:
            outSheet: sheet of workbook
            colIndex: The cell column (zero indexed).
            rowIndex: The cell row (zero indexed).

        Returns:
            cell of workbook sheet

        """

        # HACK: Extract the internal xlwt cell representation.
        row = outSheet._Worksheet__rows.get(rowIndex)
        if not row:
            return None
        cell = row._Row__cells.get(colIndex)
        return cell


###########################################################################
#
# CMPAK reconciliation reports generating.
#
###########################################################################

def main(argv):
    # HACK: Input arguments check
    # if len(argv) != 2:
    #     print("[!] Two arguments are required! Template workbook as well as generating workbook")
    #     sys.exit(1)
    # if not os.path.exists(argv[0]):
    #     print("[!] file not exists!!")
    #     sys.exit(1)
    # if not any(os.path.basename(argv[0]).endswith('xls'), os.path.basename(argv[0]).endswith('xlsx')):
    #     print("[!] excel workbook file please!!!")
    # END HACK

    # HACK: logic check

    template = './template/Reconciliation.xls'
    report = 'C:\\Users\\Administrator\\Desktop\\output.xls'

    workbook = xlrw(template, report)

    workbook.get_workbook().save(report)

    # outSheet = workbook.sheet_by_index(0)
    # outSheet1 = workbook.sheet_by_name("GPRS")
    # workbook.set_out_cell(outSheet, 0, 0, '!!!DONE!!!!!!!!!!')

    # number format & border must be hand over, as well as write format for cell
    # workbook.set_out_cell(outSheet, 2, 1, 121112)

    # END HACK

if __name__ == '__main__':
    main(sys.argv[1:])

