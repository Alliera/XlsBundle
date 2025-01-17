import xlrd
import json
import sys
import os
import argparse
import datetime

# For more information about cell types see documentation http://www.lexicon.net/sjmachin/xlrd.html#xlrd.Cell-class
CELL_TYPE_XLDATE = 3
CELL_TYPE_TEXT = 1


def run(argv):
    parser = argparse.ArgumentParser()
    parser.add_argument('--size')
    parser.add_argument('--start')
    parser.add_argument('--action')
    parser.add_argument('--file')
    parser.add_argument('--max-empty-rows', dest="max_empty_rows")
    parser.add_argument('--with-empty-rows')
    args = parser.parse_args()
    with_empty_rows = args.with_empty_rows

    if not os.path.isfile(args.file):
        print("File does not exist")
        sys.exit(1)

    workbook = xlrd.open_workbook(args.file)
    sheet = workbook.sheet_by_index(0)

    if args.action == "count":
        max_empty_rows = int(args.max_empty_rows)
        total_rows_count = 0
        empty_rows_count = 0
        for row_number in range(sheet.nrows):
            row = sheet.row_values(row_number)
            row_empty = True
            for cell in row:
                if cell not in (u'', ''):
                    row_empty = False
                    empty_rows_count = 0
                    break

            if not row_empty or with_empty_rows:
                total_rows_count += 1

            if row_empty:
                empty_rows_count += 1

            if max_empty_rows < empty_rows_count:
                if with_empty_rows:
                    total_rows_count = total_rows_count - empty_rows_count  # cut off trailing empty rows
                break

        print(total_rows_count)

    elif args.action == "read":
        rows = []
        num_cols = sheet.ncols
        row_process_number = 0
        for row_idx in range(int(args.start) - 1, sheet.nrows):  # Iterate through rows
            row_process_number += 1
            current_row_read = []
            row_empty = True
            for col_idx in range(0, num_cols):  # Iterate through columns
                cell_obj = sheet.cell(row_idx, col_idx)  # Get cell object by row, col
                if cell_obj.value not in (u'', ''):
                    row_empty = False

                if cell_obj.ctype == CELL_TYPE_XLDATE:
                    current_row_read.append(
                        datetime.datetime(*xlrd.xldate_as_tuple(cell_obj.value, workbook.datemode)).isoformat())
                else:
                    current_row_read.append(cell_obj.value)

            if not row_empty or with_empty_rows:
                rows.append(current_row_read)

            if row_process_number >= int(args.size):
                break

        print(json.dumps(rows))

    else:
        print("Unknown command")
        sys.exit(1)


if __name__ == "__main__":
    run(sys.argv[1:])
