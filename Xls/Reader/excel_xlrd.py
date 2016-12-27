import xlrd
import json
import sys
import os
import argparse
import datetime


# For more information about cell types see documentation http://www.lexicon.net/sjmachin/xlrd.html#xlrd.Cell-class
CELL_TYPE_XLDATE = 3


def run(argv):
    parser = argparse.ArgumentParser()
    parser.add_argument('--size')
    parser.add_argument('--start')
    parser.add_argument('--action')
    parser.add_argument('--file')
    parser.add_argument('--max-empty-rows', dest="max_empty_rows")
    args = parser.parse_args()

    if False == os.path.isfile(args.file):
        print("File does not exist")
        sys.exit(1)

    workbook = xlrd.open_workbook(args.file)
    sheet = workbook.sheet_by_index(0)
    if args.action == "count":
        max_count_empty_rows = int(args.max_empty_rows)
        rows_count = 0
        empty_rows_count = 0
        for rownum in range(sheet.nrows):
            row = sheet.row_values(rownum)
            current_row_read = []
            for cell in row:
                value = cell
                if value is "":
                    continue
                else:
                    current_row_read.append(True)
                    empty_rows_count = 0
                    break
            if not current_row_read:
                empty_rows_count += 1
            else:
                rows_count += 1
            if max_count_empty_rows < empty_rows_count:
                break
        print(rows_count)

    elif args.action == "read":
        rows = []
        num_cols = sheet.ncols
        for row_idx in range(int(args.start) - 1, sheet.nrows):  # Iterate through rows
            current_row_read = []
            for col_idx in range(0, num_cols):  # Iterate through columns
                cell_obj = sheet.cell(row_idx, col_idx)  # Get cell object by row, col
                if cell_obj.ctype == CELL_TYPE_XLDATE:
                    current_row_read.append(datetime.datetime(*xlrd.xldate_as_tuple(cell_obj.value, workbook.datemode)).isoformat())
                else:
                    current_row_read.append(cell_obj.value)
            rows.append(current_row_read)
            if len(rows) >= int(args.size):
                break
        print(json.dumps(rows))

    else:
        print("Unknown command")
        sys.exit(1)


if __name__ == "__main__":
    run(sys.argv[1:])
