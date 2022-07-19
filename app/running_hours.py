from app.utils import *


def generateRHData(
    file_name: str,
    vessels: list,
    machineries: list,
    debugMode: bool,
    keys: list,
    _type: str,
    showExtraInfo: bool = True,
    separateExcel: bool = True,
    _sheet: any = None,
):
    try:
        # Preparation
        error = False
        path = "src/" + file_name
        if showExtraInfo:
            console.print("\n\n📑 " + file_name, style="white", highlight=False)
        data = pd.read_excel(path, sheet_name=None, index_col=None, header=None)
        if separateExcel:
            book = Workbook()
            sheet = book.active
            sheet.append(rh_header)

        # Encoding
        vessel_id = str(data[keys[12]].iloc[0, 2])
        vessel = getVessel(
            vessel_id,
            "running_hours",
            file_name,
            vessels,
        )

        if isEmpty(vessel):
            error = True
            createLog(
                file_name,
                vessel,
                "running_hours",
                "❌ Vessel is undefined, failed to encode data. "
                + "(File: "
                + file_name
                + ", Sheet: "
                + str(key)
                + ")",
            )

        else:
            in_key = keys
            if showExtraInfo:
                in_key = track(
                    keys, description="🟢 [bold green]Processing[/bold green]"
                )

            for key in in_key:
                if key not in not_included:

                    machinery_id = str(data[key].iloc[2, 5])
                    machinery = getMachinery(
                        machinery_id,
                        key,
                        "running_hours",
                        file_name,
                        machineries,
                        vessel,
                    )

                    if (
                        machinery == "Ballast Water Management System"
                        and _type == "engine"
                    ):
                        machinery = "Ballast Water Treatment System"

                    if isEmpty(machinery):
                        error = True
                        createLog(
                            file_name,
                            vessel,
                            "running_hours",
                            "❌ Machinery code is undefined "
                            + "(File: "
                            + file_name
                            + ", Sheet: "
                            + str(key)
                            + ")",
                        )
                    else:
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
                                    vessel,
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
                                    vessel,
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
                        if separateExcel:
                            sheet.append(rowData)
                        else:
                            _sheet.append(rowData)

                        global rh_count
                        rh_count += 1

            if separateExcel:
                creation_folder = "./res/" + vessel + "/running_hours/"
                _filename = (
                    str(file_name[: len(file_name) - 5]).strip()
                    + " (Running Hours)"
                    + ".xlsx"
                )
                saveExcelFile(book, _filename, creation_folder)

            if error and not debugMode and showExtraInfo:
                console.print(
                    "❌ Error(s) found, refer to the log folder for more information.",
                    style="danger",
                    highlight=False,
                )

            if showExtraInfo:
                console.print("📥 Completed", style="info", highlight=False)

        return True

    except Exception as e:
        console.print("❌ Error: " + str(e), style="danger", highlight=False)


def running_hours(debugMode: bool):
    refresh = True
    processDone = isError = isExceptionError = False
    while True:
        try:
            resetCleanedList()

            if refresh:
                srcData = processSrc("🏃 [yellow]Running Hours[/yellow]", True)
                refresh = False

            header()
            console.print("", srcData["table"], "\n")

            if isError:
                console.print(
                    "❌ Error: " + "You have selected an invalid option.",
                    style="danger",
                    highlight=False,
                )

            if isExceptionError and debugMode:
                console.print("❌ " + exceptionMsg, style="danger", highlight=False)

            if debugMode:
                console.print("🛠️ Debug Mode: On", style="success", highlight=False)

            user_input = Prompt.ask(
                "[blink yellow]👉 Select an option[/blink yellow]",
            )

            vessels = getVessels()
            machineries = getMachineries()

            if user_input.upper() == "A":
                for _file in srcData["files"]:
                    processDone = generateRHData(
                        _file["excelFile"],
                        vessels,
                        machineries,
                        debugMode,
                        _file["keys"],
                        _file["type"],
                    )
            elif user_input.upper() == "D":
                for _file in srcData["files"]:
                    if _file["type"] == "deck":
                        processDone = generateRHData(
                            _file["excelFile"],
                            vessels,
                            machineries,
                            debugMode,
                            _file["keys"],
                            _file["type"],
                        )
            elif user_input.upper() == "E":
                for _file in srcData["files"]:
                    if _file["type"] == "engine":
                        processDone = generateRHData(
                            _file["excelFile"],
                            vessels,
                            machineries,
                            debugMode,
                            _file["keys"],
                            _file["type"],
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
                    vessels,
                    machineries,
                    debugMode,
                    srcData["files"][int(user_input) - 1]["keys"],
                    srcData["files"][int(user_input) - 1]["type"],
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


def running_hours_all(
    srcData: dict, vessels: list, machineries: list, debugMode: bool, folder_name: str
):
    try:
        resetCleanedList()

        console.print("\n\n🏃 Running Hours", highlight=False)

        book = Workbook()
        sheet = book.active
        sheet.append(rh_header)

        for _file in track(
            srcData["files"],
            description="🟢 [bold green]Processing [/bold green]",
        ):
            _ = generateRHData(
                _file["excelFile"],
                vessels,
                machineries,
                debugMode,
                _file["keys"],
                _file["type"],
                False,
                False,
                sheet,
            )

        creation_folder = "./res/AIO/" + folder_name + "/running_hours/"
        _filename = folder_name + " (Running Hours)" + ".xlsx"
        saveExcelFile(book, _filename, creation_folder)

        global rh_count
        if rh_count > 1:
            console.print(
                "🔵 Total Data Encoded: " + str(rh_count) + " Rows",
                style="bold cyan",
                highlight=False,
            )
        else:
            console.print(
                "🔵 Total Data Encoded: " + str(rh_count) + " Row",
                style="bold cyan",
                highlight=False,
            )

        value = getMinVal(rh_count)

        if value > 1:
            console.print(
                "🟣 Min Rows Per Excel: " + str(value) + " Rows",
                style="bold magenta",
                highlight=False,
            )
        else:
            console.print(
                "🟣 Min Rows Per Excel: " + str(value) + " Row",
                style="bold magenta",
                highlight=False,
            )

        excel_count = 1
        global base
        if rh_count >= base:
            excel_count = splitAIO(creation_folder, _filename, "Running Hours", value)

        if excel_count > 1:
            console.print(
                "🟡 Total File Created: " + str(excel_count) + " Files",
                style="bold yellow",
                highlight=False,
            )
        else:
            console.print(
                "🟡 Total File Created: " + str(excel_count) + " File",
                style="bold yellow",
                highlight=False,
            )

        console.print("📥 Completed", style="info", highlight=False)
    except Exception as e:
        if debugMode:
            logger.exception(e, stack_info=True)
