from os.path import exists
from app.definitions import *
from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime
import pandas as pd
import os
import re


def getMachinery(machinery_code: str, key: str, mode: str, file_name: str):
    try:
        machinery_code = machinery_code.rstrip()

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


def getIntervals(mode: int):
    try:
        intervals = []

        path = "./data/interval_list.xlsx"
        interval_list = pd.read_excel(path)

        last = "END"

        i = 0
        while (not pd.isna(interval_list.iloc[i, mode])) and (
            interval_list.iloc[i, mode] != last
        ):
            intervals.append(str(interval_list.iloc[i, mode]).rstrip())
            i += 1
            if interval_list.iloc[i, mode] == last:
                break

        return intervals
    except Exception as e:
        print("Error: " + str(e))


def getInterval(interval_id: str, interval_ids: list, interval_names: list):
    try:
        interval_id = interval_id.rstrip()
        idx = interval_ids.index(interval_id)

        return interval_names[idx]
    except ValueError:
        print("Error: " + str(interval_id) + " is not a valid interval id.")
        return ""
    except Exception as e:
        print("Error: " + str(e))


def processSrc(mode: str):
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
