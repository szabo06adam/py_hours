import sys
import os
import glob
import re
import argparse
import math
import openpyxl


quarters = [[1,2,3], [4,5,6], [7,8,9], [10,11,12]]
year_month = re.compile("(\d{4}\.)(\d{2})")

hours_in_day = 8
smart_files = True
diff_pos = False
diff_neg = False
out_file = None

sheet_to_book = {}
book_to_path = {}
path_to_diffs = {}

def MonthsInQuarter(quarter):
    return quarters[quarter]

def ToOutput(output_str):
    print(output_str)
    if out_file is not None:
        out_file.write(output_str + "\n")

def TryGetFile(path):
    if os.path.isfile(path):
        return path
    else:
        return None

def SmartFilesDir(path):
    dir_files = glob.glob(os.path.join(path, "*.xlsx"))
    dir_files.sort(reverse=True)
    if len(dir_files) < 1:
        return []
    return SmartFilesFile(dir_files[0])

def SmartFilesFile(path):
    paths = []
    # find out what month is it with regex
    matches = re.finditer(year_month, path)
    for match in matches:
        val = int(match.group(2))
        quarter = (val - 1) // 3
        months = MonthsInQuarter(quarter)
        for month in months:
            file = re.sub(year_month, f"\\g<1>{month:02d}", path)
            file_path = TryGetFile(file)
            if file_path is not None:
                paths.append(file_path)
    return paths

def SmartFiles(path):
    paths = []
    if os.path.isdir(path):
        paths = SmartFilesDir(path)
    elif os.path.isfile(path):
        paths = SmartFilesFile(path)
    else:
        print("ERR: " + path + " is not a valid file or directory!")
    return paths

def GetFiles(file_list):
    excels = []
    if smart_files and len(file_list) < 2:
        path = file_list[0]
        files = SmartFiles(path)
        for file in files:
            excels.append(file)
    else:
        for path in file_list:
            if os.path.isfile(path):
                excels.append(path)
            else:
                print("ERR: " + path + " is not a valid file or directory!")
    excels.sort()
    return excels

def OpenBooks(paths):
    global book_to_path
    global path_to_diffs

    workbooks = []
    for path in paths:
        book = openpyxl.load_workbook(path, read_only=True, data_only=True)
        workbooks.append(book)
        book_to_path[book] = path
        path_to_diffs[path] = []
    return workbooks

def OpenSheets(books):
    global sheet_to_book

    sheets = []
    for i in range(0, len(books)):
        sheet = books[i].worksheets[0]
        sheets.append(sheet)
        sheet_to_book[sheet] = books[i]
    return sheets

def isRegularworkFinished(sheet, row):
    return sheet['E' + str(row)].value is not None

def IsOvertimeWorkFinished(sheet, row):
    return sheet['H' + str(row)].value is not None

def IsWorkingToday(sheet, row):
    return sheet['C' + str(row)].value is None

def SumSheet(sheet):
    hour_sum = 0
    lastday = sheet['J9'].value[::-1][1:3][::-1]
    firstday_row = 14
    lastday_row  = firstday_row + int(lastday) - 1
    for i in range(firstday_row, lastday_row + 1):
        if sheet['D' + str(i)].protection.locked:   #not workday
            if  IsOvertimeWorkFinished(sheet, i):
                cell_val = sheet['T' + str(i)].value
                tmp_sum = cell_val
                if (diff_pos and tmp_sum > 0) or (diff_neg and tmp_sum < 0):
                    path_to_diffs[book_to_path[sheet_to_book[sheet]]].append((sheet['B' + str(i)].value, tmp_sum))
                hour_sum += tmp_sum
        else:                                       #regular workday
            if isRegularworkFinished(sheet, i) or IsOvertimeWorkFinished(sheet, i) or not IsWorkingToday(sheet, i):
                cell_val = sheet['T' + str(i)].value
                tmp_sum = cell_val - hours_in_day
                if (diff_pos and tmp_sum > 0) or (diff_neg and tmp_sum < 0):
                    path_to_diffs[book_to_path[sheet_to_book[sheet]]].append((sheet['B' + str(i)].value, tmp_sum))
                hour_sum += tmp_sum
    return hour_sum

def SumHours(sheets):
    sums = []
    for sheet in sheets:
        sums.append(SumSheet(sheet))

    hour_sum = 0
    for s in sums:
        hour_sum += s
    return hour_sum

def PrintDiffs(file):
    diffs = path_to_diffs[file]
    if len(diffs) < 1:
        return
    ToOutput(file + ":")
    for day, hour in diffs:
        ToOutput(str(day) + "\t" + str(hour))

# Argument parsing:
def SetSmartFiles(val):
    global smart_files
    smart_files = val

def SetShowDiff(val):
    if val == 'positive' or val == 'both':
        global diff_pos
        diff_pos = True
    if val == 'negative' or val == 'both':
        global diff_neg
        diff_neg = True
    # convert to pos and neg bools and use them to print affected dates

def SetOutput(val):
    global out_file
    out_file = open(val, "w", encoding='utf-8')
    # output dates (if -D) and end sum to given file
    # decide what format to use (CSV?)

def SetHours(val):
    global hours_in_day
    hours_in_day = val

def main():
    parser = argparse.ArgumentParser(
        prog='hours.py',
        formatter_class=argparse.RawTextHelpFormatter,
        description='Calculates the difference between required work hours and hours worked in a quarter.',
        # epilog='goodbye'
        )
    parser.add_argument('-O',
                        '--one-file',
                        help='disable automatic quarter recognition if only passed one file',
                        action='store_true',
    )
    parser.add_argument('-D',
                        '--show-difference',
                        help='print dates where hours worked differ from required hours\n(only shows positive/negative difference days if specified)',
                        nargs='?',
                        choices=['positive', 'negative', 'both'],
                        const='both',
                        required=False
    )
    parser.add_argument('-o',
                        '--output',
                        help='path to an output file',
                        nargs='?',
                        const='output.txt',
                        required=False
    )
    parser.add_argument('-p',
                        '--part-time',
                        help='how many hours a workday should be',
                        nargs='?',
                        default=8,
                        type=int,
                        action='store'
    )
    parser.add_argument('-H',
                        '--hour-min',
                        help='format output into hh:mm format instead of decimal time',
                        action='store_true'
    )
    parser.add_argument('path_to_hours_xlsx',
                        help="path to a file containing the work hour information",
                        nargs='*',
                        default=[os.getcwd()]
    )

    args = parser.parse_args()

    SetSmartFiles(not args.one_file)

    if args.show_difference is not None:
        SetShowDiff(args.show_difference)

    if args.output is not None:
        SetOutput(args.output)

    SetHours(args.part_time)

    filepaths = GetFiles(args.path_to_hours_xlsx)
    workbooks = OpenBooks(filepaths)
    sheets = OpenSheets(workbooks)
    hours_sum = SumHours(sheets)

    for file in filepaths:
        PrintDiffs(file)

    print("========================")
    if args.hour_min:
        h = math.floor(hours_sum)
        m = int((hours_sum - h) * 60)
        hhmm = str(h) + ":" + str(m)
        ToOutput("\n" + hhmm)
    else:
        ToOutput("\n" + str(hours_sum))

    if out_file is not None:
        out_file.close()


if __name__ == "__main__":
    main()
