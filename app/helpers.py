from os.path import exists
import time
from app.definitions import *
from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime
import pandas as pd
import os
import re


def createBin(file_name: str, mode: str, key: str, desc: str):
    creation_name = "/" + file_name
    creation_path = "./bin/" + mode

    if not os.path.exists(creation_path):
        os.makedirs(creation_path)

    if not exists(creation_path + creation_name):
        writer = pd.ExcelWriter(creation_path + creation_name, engine="xlsxwriter")
        writer.save()
        book = load_workbook(creation_path + creation_name)
        sheet = book.active
        sheet.append(bin_header)
        book.save(creation_path + creation_name)

    book = load_workbook(creation_path + creation_name)
    sheet = book.active

    rowData = (key, desc)
    sheet.append(rowData)
    book.save(creation_path + creation_name)

    print(desc)


def getMachineries():
    try:
        machineries: list = []

        path = "./data/name_list.xlsx"
        mach_list = pd.read_excel(path)

        i = 0
        while not pd.isna(mach_list.iloc[i, 1]) and mach_list.iloc[i, 1] != "END":
            machineries.append([str(mach_list.iloc[i, 0]), str(mach_list.iloc[i, 1])])
            i += 1

        return machineries
    except Exception as e:
        print("❌ Error: " + str(e))


def getMachinery(
    machinery_id: str,
    key: str,
    mode: str,
    file_name: str,
    machineries: list,
):
    try:
        if not pd.isna(machinery_id) or machinery_id != "":
            machinery_id = machinery_id.strip()

        for machinery in machineries:
            if machinery[1] == machinery_id or machinery[1] == key:
                return str(machinery[0])

        createBin(
            file_name,
            mode,
            key,
            "⚠️ Warning: No machinery ( " + machinery_id + " ) found for " + key,
        )

        return "N/A"

    except Exception as e:
        print("❌ Error: " + str(e) + " (" + key + ": " + machinery_id + ")")


def getCodes():
    try:
        codes: list = []

        path = "./data/code_list.xlsx"
        code_list = pd.read_excel(path)

        i = 0
        while not pd.isna(code_list.iloc[i, 1]) and code_list.iloc[i, 1] != "END":
            codes.append([str(code_list.iloc[i, 0]), str(code_list.iloc[i, 1])])
            i += 1

        return codes
    except Exception as e:
        print("❌ Error: " + str(e))


def getCode(
    machinery_name: str,
    key: str,
    mode: str,
    file_name: str,
    codes: list,
):
    try:
        if not pd.isna(machinery_name) or machinery_name != "":
            machinery_name = machinery_name.strip()

        for code in codes:
            if code[1] == machinery_name or code[1] == key:
                return str(code[0])

        createBin(
            file_name,
            mode,
            key,
            "⚠️ Warning: No machinery code ( " + machinery_name + " ) found for " + key,
        )

        return "N/A"
    except Exception as e:
        print("❌ Error: " + str(e) + " (" + key + ": " + machinery_name + ")")


def getIntervals():
    try:
        intervals: list = []

        path = "./data/interval_list.xlsx"
        interval_list = pd.read_excel(path)

        i = 0
        while (
            not pd.isna(interval_list.iloc[i, 1]) and interval_list.iloc[i, 1] != "END"
        ):
            intervals.append(
                [str(interval_list.iloc[i, 0]), str(interval_list.iloc[i, 1])]
            )
            i += 1

        return intervals
    except Exception as e:
        print("❌ Error: " + str(e))


def getInterval(
    interval_id: str,
    key: str,
    mode: str,
    file_name: str,
    intervals: list,
):
    try:
        if not pd.isna(interval_id) or interval_id != "":
            interval_id = interval_id.strip()

        for interval in intervals:
            if interval[1] == interval_id or interval[1] == key:
                return str(interval[0])

        createBin(
            file_name,
            mode,
            key,
            "⚠️ Warning: No interval ( " + interval_id + " ) found for " + key,
        )

        return "N/A"
    except Exception as e:
        print("❌ Error: " + str(e) + " (" + key + ": " + interval_id + ")")


def has_numbers(inputString):
    return bool(re.search(r"\d", inputString))


def header(title: str):
    print(
        r"""
____              _____                     _           
|  _ \ _   _      | ____|_ __   ___ ___   __| | ___ _ __ 
| |_) | | | |_____|  _| | '_ \ / __/ _ \ / _` |/ _ \ '__|
|  __/| |_| |_____| |___| | | | (_| (_) | (_| |  __/ |   
|_|    \__, |     |_____|_| |_|\___\___/ \__,_|\___|_|   
        |___/                                             
            """
    )

    print(title + "\n")


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
            print("No data found in src directory.")
            time.sleep(5)
            exit()

        print("A - All")

        return files
    except Exception as e:
        print("❌ Error: " + str(e))


def isEmpty(data: str):
    if (data == "") or (data == " ") or (pd.isna(data)) or (data == "nan"):
        return True
    else:
        return False


def isValid(data: str):
    if (
        (data == "")
        or (data == " ")
        or (pd.isna(data))
        or (data == "Note:") 
        or (data == "nan")
        or not (has_numbers(data))
    ):
        return False
    else:
        return True


def exitApp():
    isContinue = input("\n🟢 Input 1 to continue: ")
    if isContinue == "1":
        return False
    else:
        return True
