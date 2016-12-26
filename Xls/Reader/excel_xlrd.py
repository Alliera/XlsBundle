import xlrd
import json
import sys
import os
import argparse


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
        reached_end = False
        rows = []
        while len(rows) < int(args.size) and reached_end == False:
            try:
                rows.append(sheet.row_values(int(args.start) + len(rows) - 1))
            except IndexError:
                reached_end = True
        print(json.dumps(rows))

    else:
        print("Unknown command")
        sys.exit(1)


if __name__ == "__main__":
    run(sys.argv[1:])
