from os.path import exists
from app.definitions import *
from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime
import pandas as pd
import os
import re


def generateMainData(file_name):
    try:
        path = "src/" + file_name
        print("Excel File: " + file_name)

        # Read the data
        data = pd.read_excel(path, sheet_name=None, index_col=None, header=None)

        # Get the keys
        xl = pd.ExcelFile(path)
        keys = xl.sheet_names

        # Iterate through the sheets
        for key in keys:
            if key not in not_included:
                print(str(key).rstrip())

                # Vessel Name
                vessel = str(data[key].iloc[0, 2])

                # Default Machinery Name: machinery = data[key].iloc[2, 2]
                # Machinery Name using the machinery code
                machinery = getMachinery(
                    str(data[key].iloc[2, 5]).rstrip(), key, "main", file_name
                )

                if (not pd.isna(machinery)) and (not pd.isna(vessel)):
                    # Start traversing the data on row 7
                    row = 7
                    is_Valid = True

                    # Prepare the sheets
                    book = Workbook()
                    sheet = book.active

                    sheet.append(main_header)

                    while is_Valid:

                        rowData = (
                            vessel.rstrip(),
                            machinery.rstrip(),
                        )

                        for col in range(7):
                            d = data[key].iloc[row, col]

                            if (pd.isna(d)) and (col == 0):
                                is_Valid = False
                                break

                            if pd.isna(d):
                                d = " "

                            if (col == 3) and not (re.search("[a-zA-Z]", str(d))):
                                d = str(d) + " Hours"

                            if ((col == 4) or (col == 5)) and isinstance(d, datetime):
                                d = d.strftime("%d-%b-%y")
                            else:
                                d = re.sub("\\s+", " ", str(d))

                            tempTuple = (d.rstrip(),)
                            rowData += tempTuple

                        if is_Valid:
                            sheet.append(rowData)
                            row += 1

                    create_name = file_name[: len(file_name) - 4]
                    creation_folder = "./res/main/" + create_name
                    if not os.path.exists(creation_folder):
                        os.makedirs(creation_folder)
                    book.save(creation_folder + "/" + key + ".xlsx")

        print("Done...")
    except Exception as e:
        print("Error: " + str(e))


def generateRHData(file_name):
    try:
        path = "src/" + file_name
        print("Excel File: " + file_name)

        # Read the data
        data = pd.read_excel(path, sheet_name=None, index_col=None, header=None)

        # Get the keys
        xl = pd.ExcelFile(path)
        keys = xl.sheet_names

        # Prepare the sheets
        book = Workbook()
        sheet = book.active

        # Append the header
        sheet.append(rh_header)

        # Iterate through the sheets
        for key in keys:
            if key not in not_included:
                print(str(key).rstrip())

                # Vessel Name
                vessel = str(data[key].iloc[0, 2])

                # Machinery Name
                machinery = getMachinery(
                    str(data[key].iloc[2, 5]), key, "sub", file_name
                )

                # Running Hours
                running_hours = str(data[key].iloc[3, 5])

                # Updated At
                if not pd.isna(data[key].iloc[4, 5]):
                    updating_date = data[key].iloc[4, 5].strftime("%d-%b-%y")
                else:
                    updating_date = " "

                rowData = (
                    vessel.rstrip(),
                    machinery.rstrip(),
                    running_hours.rstrip(),
                    updating_date.rstrip(),
                )
                sheet.append(rowData)

        create_name = file_name[: len(file_name) - 4]
        creation_folder = "./res/running_hours/" + create_name
        if not os.path.exists(creation_folder):
            os.makedirs(creation_folder)
        book.save(creation_folder + "/" + file_name)

        print("Done...")
    except Exception as e:
        print("Error: " + str(e) + "\n")


def generateIntervalData(file_name):
    try:
        path = "src/" + file_name
        print("Excel File: " + file_name)

        # Read the data
        data = pd.read_excel(path, sheet_name=None, index_col=None, header=None)

        # Get the keys
        xl = pd.ExcelFile(path)
        keys = xl.sheet_names

        # Prepare the sheets
        book = Workbook()
        sheet = book.active

        # Append the header
        sheet.append(interval_header)

        # Array of intervals
        intervals = []

        # Iterate through the sheets
        for key in keys:
            if key not in not_included:
                print(str(key).rstrip())

                # Vessel Name
                vessel = str(data[key].iloc[0, 2])

                # Machinery Name using the machinery code
                machinery = str(
                    getMachinery(
                        str(data[key].iloc[2, 5]).rstrip(), key, "main", file_name
                    )
                )

                if (not pd.isna(machinery)) and (not pd.isna(vessel)):
                    # Start traversing the data on row 7
                    row = 7
                    is_Valid = True

                    while is_Valid:

                        # Interval
                        interval = str(data[key].iloc[row, 3])

                        if pd.isna(data[key].iloc[row, 0]):
                            is_Valid = False
                            break

                        if pd.isna(data[key].iloc[row, 3]):
                            interval = " "

                        # Check if the interval is hours
                        if not re.search("[a-zA-Z]", interval) and interval != " ":
                            interval = interval + " Hours"

                        # If the interval is not yet written then it will be written
                        # Otherwise, not
                        if interval not in intervals:
                            intervals.append(interval)
                            rowData = (vessel.rstrip(), interval.rstrip())
                            sheet.append(rowData)

                        row += 1

        create_name = file_name[: len(file_name) - 4]
        creation_folder = "./res/interval/" + create_name
        if not os.path.exists(creation_folder):
            os.makedirs(creation_folder)
        book.save(creation_folder + "/" + file_name)

    except Exception as e:
        print("Error: " + str(e))


def getMachinery(machinery_code, key, mode, file_name):
    try:
        machinery_code = machinery_code.rstrip()

        if not os.path.exists("./data"):
            os.makedirs("./data")

        if machinery_code == "nan" or machinery_code == "none":
            machinery_code = key

        path = "./data/gen_mach_list.xlsx"
        mach_list = pd.read_excel(path)

        last = "END"

        i = 0
        while (not pd.isna(mach_list.iloc[i, 1])) and (
            mach_list.iloc[i, 1] != machinery_code
        ):
            i += 1
            if mach_list.iloc[i, 1] == last:
                break

        if not pd.isna(mach_list.iloc[i, 1]) and (
            mach_list.iloc[i, 1] == machinery_code
        ):
            return str(mach_list.iloc[i, 0])
        else:
            creation_name = "/" + file_name
            creation_path = "./bin/" + mode

            if not os.path.exists(creation_path):
                os.makedirs(creation_path)

            if not exists(creation_path + creation_name):
                writer = pd.ExcelWriter(
                    creation_path + creation_name, engine="xlsxwriter"
                )
                writer.save()
                book = load_workbook(creation_path + creation_name)
                sheet = book.active
                sheet.append(bin_header)
                book.save(creation_path + creation_name)

            book = load_workbook(creation_path + creation_name)
            sheet = book.active

            rowData = (key, machinery_code)
            sheet.append(rowData)
            book.save(creation_path + creation_name)

            print(
                "\nWarning: No machinery code found for "
                + key
                + " ( "
                + machinery_code
                + " )\n"
            )
            return "N/A"

    except Exception as e:
        print("Error: " + str(e) + " (" + key + ": " + machinery_code + ")" + "\n")


def processSrc(mode):
    try:
        mode_path = "./res/" + mode
        if not os.path.exists(mode_path):
            os.makedirs(mode_path)

        if not os.path.exists("./src"):
            os.makedirs("./src")

        files = []
        i = 0
        for excel in os.listdir("./src"):
            if excel.endswith(".xlsx"):
                files.append(excel)
                print(i, "-", excel)
                i += 1

        if len(files) == 0:
            print("No such data found in src directory.")
            return []

        print("A - All")

        return files
    except Exception as e:
        print("Error: " + str(e))


def exitApp():
    isContinue = input("Input 1 to continue: ")
    if isContinue == "1":
        return False
    else:
        return True
