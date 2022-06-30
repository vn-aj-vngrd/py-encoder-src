from app.helpers import *
import os


def mainFunction():
    try:
        while True:
            files = processSrc("main")
            if len(files) == 0:
                break

            file_key = input("\nSelect an option: ")

            if file_key != "A":
                file_name = files[int(file_key)]
                generateMainData(file_name)

            else:
                for _file in files:
                    generateMainData(_file)

            if exitApp():
                break

    except Exception as e:
        print("Error: " + str(e))


def rhFunction():
    try:
        while True:
            files = processSrc("running_hours")
            if len(files) == 0:
                break

            file_key = input("\nSelect an option: ")

            if file_key != "A":
                file_name = files[int(file_key)]
                generateRHData(file_name)
            else:
                for _file in files:
                    generateRHData(_file)

            if exitApp():
                break

    except Exception as e:
        print("Error: " + str(e))


def intervalFunction():
    try:
        while True:
            files = processSrc("interval")
            if len(files) == 0:
                break

            file_key = input("\nSelect an option: ")

            if file_key != "A":
                file_name = files[int(file_key)]
                generateIntervalData(file_name)
            else:
                for _file in files:
                    generateIntervalData(_file)

            if exitApp():
                break

    except Exception as e:
        print("Error: " + str(e))
