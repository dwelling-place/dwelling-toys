import xlrd
import sys
import datetime


def castcell(book, cell):
    if cell.ctype == xlrd.XL_CELL_DATE:
        return datetime.datetime(xldate_as_tuple(cell.value, book.datemode))
    else:
        return cell.value

book = open_workbook(sys.argv[1])

for sheet in book.sheets():
    data = [[None] * sheet.ncols] * sheet.nrows
    for rown in range(sheet.nrows):
        for coln in range(sheet.ncols):
            data[rown][coln] = castcell(book, sheet.cell(rown, coln))

    period = row[2][0]
    del data[0:4]  # Sheet header is 4 rows

    # Row 0 is the property headers
    # Row 1 is Actual/Budget
    # Row 2... is data

    props = {
        ...
    }
