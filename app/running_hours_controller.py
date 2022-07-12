from app.utils import *


def generateRHData(file_name: str, machineries: list, debugMode: bool, keys: list):
    try:
        path = "src/" + file_name
        console.print("\n\n📑 " + file_name, style="white", highlight=False)

        data = pd.read_excel(path, sheet_name=None, index_col=None, header=None)

        book = Workbook()
        sheet = book.active
        sheet.append(rh_header)

        vessel = str(data[keys[12]].iloc[0, 2])

        error = False

        for key in track(keys, description="🟢 [bold green]Processing[/bold green]"):
            if key not in not_included:
                machinery_id = str(data[key].iloc[2, 5])

                machinery = getMachinery(
                    machinery_id,
                    key,
                    "running_hours",
                    file_name,
                    machineries,
                )

                if not isEmpty(vessel) and not isEmpty(machinery):

                    # Running Hours
                    running_hours = data[key].iloc[3, 5]
                    if isEmpty(running_hours):
                        running_hours = "0"
                    else:
                        if str(running_hours) == "10.737.2":
                            running_hours = 10737.2
                        elif not isFloat(str(running_hours)):
                            error = True

                            createLog(
                                file_name,
                                "running_hours",
                                '❌ Running Hours "'
                                + str(running_hours)
                                + '" is invalid '
                                + "(File: "
                                + str(file_name)
                                + ", Sheet: "
                                + str(key)
                                + ", Machinery: "
                                + str(machinery)
                                + ")",
                            )
                            running_hours = "0"

                    # Updating Date
                    updating_date = data[key].iloc[4, 5]
                    if isEmpty(updating_date):
                        updating_date = ""
                    else:
                        if isinstance(updating_date, datetime):
                            updating_date = updating_date.strftime("%d-%b-%y")
                        else:
                            error = True
                            createLog(
                                file_name,
                                "running_hours",
                                '❌ Updating date "'
                                + str(updating_date)
                                + '" is invalid '
                                + "(File: "
                                + file_name
                                + ", Sheet: "
                                + str(key)
                                + ", Machinery: "
                                + machinery
                                + ")",
                            )
                            updating_date = ""

                    rowData = (
                        vessel,
                        machinery,
                        str(running_hours).strip(),
                        str(updating_date).strip(),
                    )
                    sheet.append(rowData)
                else:
                    error = True
                    createLog(
                        file_name,
                        "running_hours",
                        "❌ Vessel name or machinery code is empty "
                        + "(File: "
                        + file_name
                        + ", Sheet: "
                        + str(key)
                        + ")",
                    )

        if not os.path.exists("./res/running_hours"):
            os.makedirs("./res/running_hours")

        _filename = (
            str(file_name[: len(file_name) - 5]).strip() + " (Running Hours)" + ".xlsx"
        )
        creation_folder = "./res/running_hours/"
        saveExcelFile(book, _filename, creation_folder)

        if error and not debugMode:
            console.print(
                "❌ Error(s) found, refer to the log folder for more information.",
                style="danger",
                highlight=False,
            )

        console.print("📥 Completed", style="info")
        return True

    except Exception as e:
        console.print("❌ Error: " + str(e), style="danger")


def running_hours(debugMode: bool):
    refresh = True
    processDone = isError = isExceptionError = False
    while True:
        try:
            global cleaned_log_list
            cleaned_log_list.clear()

            if refresh:
                srcData = processSrc("🏃 [yellow]Running Hours[/yellow]", True)
                refresh = False

            header()
            console.print("", srcData["table"], "\n")

            if isError:
                console.print(
                    "❌ Error: " + "You have selected an invalid option.",
                    style="danger",
                )

            if isExceptionError and debugMode:
                console.print("❌ " + exceptionMsg, style="danger")

            if debugMode:
                console.print("🛠️ Debug Mode: On", style="success")

            user_input = Prompt.ask(
                "[blink yellow]👉 Select an option[/blink yellow]",
            )

            machineries = getMachineries()

            if user_input.upper() == "A":
                for _file in srcData["files"]:
                    processDone = generateRHData(
                        _file["excelFile"], machineries, debugMode, _file["keys"]
                    )
            elif user_input.upper() == "D":
                for _file in srcData["files"]:
                    if _file["type"] == "deck":
                        processDone = generateRHData(
                            _file["excelFile"], machineries, debugMode, _file["keys"]
                        )
            elif user_input.upper() == "E":
                for _file in srcData["files"]:
                    if _file["type"] == "engine":
                        processDone = generateRHData(
                            _file["excelFile"], machineries, debugMode, _file["keys"]
                        )
            elif user_input.upper() == "G":
                break
            elif user_input.upper() == "R":
                refresh = True
            elif (
                user_input.isdigit()
                and int(user_input) >= 1
                and int(user_input) <= len(srcData["files"])
            ):
                processDone = generateRHData(
                    srcData["files"][int(user_input) - 1]["excelFile"],
                    machineries,
                    debugMode,
                    srcData["files"][int(user_input) - 1]["keys"],
                )
            else:
                isError = True

            if processDone:
                isError = processDone = False
                if promptExitorContinue():
                    break

        except Exception as e:
            isExceptionError = True
            exceptionMsg = str(e)


def running_hours_all(srcData: dict, machineries: list, debugMode: bool):
    try:
        global cleaned_log_list
        cleaned_log_list.clear()

        console.print("[magenta]-[/magenta]" * 67)
        console.print("🏃 [yellow]Running Hours[/yellow]")
        console.print("[magenta]-[/magenta]" * 67)

        for _file in srcData["files"]:
            _ = generateRHData(
                _file["excelFile"], machineries, debugMode, _file["keys"]
            )

    except Exception as e:
        if debugMode:
            logger.exception(e, stack_info=True)
