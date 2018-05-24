import os

import openpyxl
import pandas as pd
from openpyxl.styles import Font, Border, Side, NamedStyle

'''
@author Arjun Gupta
@date 5/23/2018
@version 1.0

This program takes in two Allegro Netlists that have been converted to Excel format (copy and paste from .htm)
and parses them to accurately and rapidly find all points at both the net and the pin level where there are 
disagreements between the two netlists. The main purpose of this is to aide in the translation of schematics
from one format (in my case, OrCAD) to another (Cadence HDL). 

To do: 
    -Go straight from htm to Excel upload (aka supporting more file formats)
    -Redesign with classes in mind - limit the amount of information reuse
    -Decide on one format - either Pandas or OpenPyxl
'''


def main():
    file_path, data, xl = find_file()
    sheet1_data = data_frame_to_dict(data[0])
    sheet2_data = data_frame_to_dict(data[1])
    diff_dict_pins, diff_dict_nets = compare_sheets(sheet1_data, sheet2_data)
    wb = openpyxl.load_workbook(filename=file_path)
    if "BorderAndFont" not in wb.named_styles:
        wb.add_named_style(create_font_style())
    export(file_path, diff_dict_pins, "Compared Pin Results", wb)
    export(file_path, diff_dict_nets, "Compared Net Results", wb)


def find_file():
    subdir = "Data"
    path = os.path.normpath(os.path.join(os.getcwd(), subdir))
    files = os.listdir(path)
    print("Current files: ")
    print("-------------------------------------------------------------")
    for f in files:
        print(f)
    print("-------------------------------------------------------------")
    file = input("Enter the file you want to read from: ")
    data_frames = None
    prev_path = path
    #Excel file
    xl = None
    while data_frames is None:
        try:
            path = os.path.normpath(os.path.join(path, file))
            extension = file.split(".")
            if len(extension) > 1:
                if extension[1] == "xls" or extension[1] == "xlsx":
                    # convert xlsx to csv here
                    xl, data_frames = excel_to_dataframe(path)
                else:
                    raise IOError("Unsupported file type selected.")
            else:
                raise IOError("Improper file name entered.")
        except OSError:
            print('Unable to find file.')
            file = input("Enter the file name that you want to read from: ")
            path = prev_path
        except IOError:
            file = input("Enter the file name that you want to read from: ")
            path = prev_path
    return path, data_frames, xl


def excel_to_dataframe(path):
    # Path verified in above file conditioning
    xl = pd.ExcelFile(path)
    sheet_names = xl.sheet_names

    # Prevents passing in a excel book with only one sheet.
    if len(sheet_names) <= 1:
        raise IOError("Improper file format.")

    print("Sheets available for parsing: ", *sheet_names, sep='\t')
    sheet1 = input("Enter first sheet to parse: ")
    sheet2 = input("Enter second sheet to parse: ")
    valid_sheets = False

    while not valid_sheets:

        if sheet1 not in sheet_names and sheet2 not in sheet_names:
            print("Both sheets were not found.\n")
            sheet1 = input("Enter first sheet to parse: ")
            sheet2 = input("Enter second to parse: ")
            continue

        elif sheet1 not in sheet_names:
            print("The first sheet was not found.\n")
            sheet1 = input("Enter first sheet to parse, second sheet is \'{}\': ".format(sheet2))
            continue

        elif sheet2 not in sheet_names:
            print("The second sheet was not found.\n")
            sheet2 = input("Enter second sheet to parse, first sheet is \'{}\': ".format(sheet1))
            continue

        elif sheet1 == sheet2:
            print("Duplicate sheet entered!\n")
            sheet2 = input("Enter a new sheet to parse, first sheet is \'{}\': ".format(sheet1))
            continue
        else:
            valid_sheets = True

    print("Sheets being compared are: {} and {}".format(sheet1, sheet2))

    sheet1_parsed = xl.parse(sheet1)
    cols = ["Net Name", "Net Pins"]
    scols = set(cols)
    svals1 = set(sheet1_parsed.columns.values.tolist())
    if not scols.issubset(svals1):
        sheet1_parsed = sheet1_parsed.drop([0])
        sheet1_parsed.columns = cols

    sheet2_parsed = xl.parse(sheet2)
    svals2 = set(sheet2_parsed.columns.values.tolist())
    if not scols.issubset(svals2):
        sheet2_parsed = sheet2_parsed.drop([0])
        sheet2_parsed.columns = cols

    return xl, [sheet1_parsed, sheet2_parsed]


def data_frame_to_dict(frame):
    frame_dict = {}
    for row in frame.values:
        if len(row) == 2:
            net_name = row[0]
            split_ref_des = row[1].split(" ")
            frame_dict[net_name] = split_ref_des
    return frame_dict


def compare_sheets(sheet1, sheet2):
    if sheet1 is None or sheet2 is None:
        raise IOError("Unable to parse due to an error.")

    if (not isinstance(sheet1, dict)) or (not isinstance(sheet2, dict)):
        raise IOError("Unable to parse something that is not a dictionary")

    diff_dict_pins = {}
    diff_dict_nets = {}
    for net_name in sheet1:
        if net_name in sheet2:
            for val1, val2 in zip(sheet1[net_name], sheet2[net_name]):
                if val1 not in sheet2[net_name]:
                    if net_name not in diff_dict_pins:
                        diff_dict_pins[net_name] = []
                    diff_dict_pins[net_name].append("[ {} v {} ]".format(val1, val2))
                    print(diff_dict_pins[net_name])
        else:
            diff_dict_nets[net_name] = sheet1[net_name]
            # print("Key: {} \t Value: {}".format(net_name, diff_dict_nets[net_name]))
    return diff_dict_pins, diff_dict_nets


def export(file_path, diff_dict, sheet_name, wb):
    if sheet_name not in wb.sheetnames:
        wb.create_sheet(title=sheet_name)
        print(wb.sheetnames)
    ws = wb[sheet_name]

    font_column = Font(name='Times New Roman',
                       size=10,
                       bold=True)

    ws['A1'] = "Net Name"
    ws['B1'] = "Net Pins"
    ws['A1'].style = "BorderAndFont"
    ws['B1'].style = "BorderAndFont"
    ws['A1'].font = font_column
    ws['B1'].font = font_column
    row = 1
    col = 1
    for key in diff_dict.keys():
        row += 1
        net_name = ws.cell(row=row, column=col)
        net_pins = ws.cell(row=row, column=col+1)
        net_name.value = key
        # print("Key: {} \t Value: {}".format(key, diff_dict[key]))

        net_pins.value = " ".join(diff_dict[key])
        net_name.style = "BorderAndFont"
        net_pins.style = "BorderAndFont"
    # file_name = ""
    # for line in file_path:
    #     drive, path = os.path.splitdrive(line)
    #     path, file_name = os.path.split(path)

    dims = {}
    for row in ws.rows:
        for cell in row:
            if cell.value:
                width = max((dims.get(cell.column, 0), len(cell.value)))
                # Padding row dimensions, look into executing a vba script
                if width >= 125:
                    factor = width // 125
                    ws.row_dimensions[cell.row].height = 15 * factor
                    continue
                dims[cell.column] = width
    for col, value in dims.items():
        ws.column_dimensions[col].width = value

    # NOTE: Program crashes when file that is being saved is open... Don't do that please
    wb.save(filename=file_path)
    wb.close()

    # TODO: Look into a solution for this? PermissionError raised but I don't know where
    # try:
    #     wb.save(filename="temp.xlsx")
    # except PermissionError:
    #     print("HERE")
    #     print(os.access("temp.xlsx", os.W_OK))
    #     while not os.access(file_path, (os.F_OK ^ os.R_OK ^ os.W_OK)):
    #         input("Waiting for user to close {}. Press Enter once done!".format(file_name))


def create_font_style():
    border_style = NamedStyle(name="BorderAndFont")
    border_style.font = Font(name='Times New Roman',
                             size=10,
                             color='000000')
    bd = Side(style='thin', color='000000')
    border_style.border = Border(left=bd, top=bd, right=bd, bottom=bd)
    border_style.alignment.wrap_text = True
    return border_style


if __name__ == '__main__':
    main()
