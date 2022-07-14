from app.utils import *


def generateUJData(
    file_name: str,
    vessels: list,
    machineries: list,
    codes: list,
    intervals: list,
    debugMode: bool,
    keys: list,
    _type: bool,
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
            sheet.append(uj_header)

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

                    machinery_id = str(data[key].iloc[2, 5]).strip()
                    machinery = getMachinery(
                        machinery_id,
                        key,
                        "update_jobs",
                        file_name,
                        machineries,
                        vessel,
                    )

                    if (
                        machinery == "Ballast Water Management System"
                        and _type == "engine"
                    ):
                        machinery = "Ballast Water Treatment System"

                    machinery_code = getCode(
                        machinery,
                        key,
                        "update_jobs",
                        file_name,
                        codes,
                        vessel,
                    )

                    if not isEmpty(machinery) and not isEmpty(machinery_code):
                        row = 7

                        while True:

                            # Code
                            code = data[key].iloc[row, 0]
                            if not isValid(code):
                                break
                            else:
                                if "-" in code:
                                    col_key = code.split("-")
                                    code = (
                                        machinery_code.rstrip()
                                        + "-"
                                        + col_key[1].lstrip()
                                    )
                                else:
                                    match = re.match(r"([a-z]+)([0-9]+)", code, re.I)
                                    if match:
                                        col_key = match.groups()

                                        code = (
                                            machinery_code.rstrip()
                                            + "-"
                                            + col_key[1].lstrip()
                                        )

                            # Name
                            name = data[key].iloc[row, 1]
                            if isEmpty(name):
                                name = ""
                            # Manual Override (--Force Fix)
                            if str(code) == "RE-009":
                                name = "EPIRB"

                            # Description
                            description = data[key].iloc[row, 2]
                            if isEmpty(description):
                                description = ""

                            # Interval
                            interval = data[key].iloc[row, 3]
                            if isEmpty(interval):
                                interval = ""
                            else:
                                if not re.search(
                                    "[a-zA-Z]", str(interval)
                                ) and not isinstance(interval, datetime):
                                    interval = str(interval) + " Hours"

                                interval = getInterval(
                                    str(interval),
                                    key,
                                    "update_jobs",
                                    file_name,
                                    intervals,
                                    str(code),
                                    vessel,
                                )

                                if isEmpty(interval):
                                    error = True
                                    interval = ""

                            # Commissioning Date
                            commissioning_date = data[key].iloc[row, 4]
                            if isEmpty(commissioning_date):
                                commissioning_date = ""
                            else:
                                if isinstance(commissioning_date, datetime):
                                    commissioning_date = commissioning_date.strftime(
                                        "%d-%b-%y"
                                    )
                                else:
                                    commissioning_date = getFormattedDate(
                                        key,
                                        code,
                                        "update_jobs",
                                        file_name,
                                        str(commissioning_date),
                                        "Commissioning date",
                                        vessel,
                                    )

                                    if isEmpty(commissioning_date):
                                        error = True

                            # Last Done Date
                            last_done_date = data[key].iloc[row, 5]
                            if isEmpty(last_done_date):
                                last_done_date = ""
                            else:
                                if isinstance(last_done_date, datetime):
                                    last_done_date = last_done_date.strftime("%d-%b-%y")
                                else:
                                    if (
                                        str(last_done_date).strip().lower()
                                        == "since new"
                                    ):
                                        last_done_date = commissioning_date
                                    else:
                                        last_done_date = getFormattedDate(
                                            key,
                                            code,
                                            "update_jobs",
                                            file_name,
                                            str(last_done_date),
                                            "Last done date",
                                            vessel,
                                        )

                                        if isEmpty(last_done_date):
                                            error = True

                            # Last Done Running Hours
                            last_done_running_hours = data[key].iloc[row, 6]
                            if isEmpty(last_done_running_hours):
                                last_done_running_hours = 0
                            else:
                                if (
                                    not isFloat(last_done_running_hours)
                                    or str(last_done_running_hours) == "True"
                                    or str(last_done_running_hours) == "False"
                                ):
                                    createLog(
                                        file_name,
                                        vessel,
                                        "update_jobs",
                                        '❌ Last done running hours "'
                                        + str(last_done_running_hours)
                                        + '" is invalid '
                                        + "(File: "
                                        + file_name
                                        + ", Sheet: "
                                        + str(key)
                                        + ", Machinery: "
                                        + str(machinery)
                                        + ", Code: "
                                        + str(code)
                                        + ")",
                                    )
                                    error = True
                                    last_done_running_hours = 0

                            # Instructions
                            instructions = data[key].iloc[row, 10]
                            if isEmpty(instructions):
                                instructions = ""

                            # Remarks
                            remarks = data[key].iloc[row, 11]
                            if isEmpty(remarks):
                                remarks = ""

                            # Insertion
                            rowData = (
                                vessel,
                                machinery,
                                code,
                                str(name).strip(),
                                re.sub("\\s+", " ", str(description).strip()),
                                str(interval).strip(),
                                str(commissioning_date).strip(),
                                str(last_done_date).strip(),
                                str(last_done_running_hours).strip(),
                                re.sub("\\s+", " ", str(instructions).strip()),
                                re.sub("\\s+", " ", str(remarks).strip()),
                            )

                            if separateExcel:
                                sheet.append(rowData)
                            else:
                                _sheet.append(rowData)
                            row += 1

                            global uj_count
                            uj_count += 1
                    else:
                        error = True
                        createLog(
                            file_name,
                            vessel,
                            "update_jobs",
                            "❌ Machinery name and code is undefined "
                            + "(File: "
                            + file_name
                            + ", Sheet: "
                            + str(key)
                            + ")",
                        )
            if separateExcel:
                creation_folder = "./res/" + vessel + "/update_jobs/"
                _filename = (
                    str(file_name[: len(file_name) - 5]).strip()
                    + " (Update Jobs)"
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
        console.print("❌ " + str(e), style="danger", highlight=False)


def update_jobs(debugMode: bool):
    refresh = True
    processDone = isError = isExceptionError = False
    while True:
        try:
            resetCleanedList()

            if refresh:
                srcData = processSrc(
                    "📝 [yellow]Update Jobs[/yellow]",
                    True,
                )
                refresh = False

            header()
            console.print("", srcData["table"], "\n")

            if isExceptionError and debugMode:
                console.print("❌ " + exceptionMsg, style="danger", highlight=False)

            if isError:
                console.print(
                    "❌ " + "You have selected an invalid option.",
                    style="danger",
                    highlight=False,
                )

            if debugMode:
                console.print("🛠️ Debug Mode: On", style="success", highlight=False)

            user_input = Prompt.ask(
                "[blink yellow]👉 Select an option[/blink yellow]",
            )

            vessels = getVessels()
            machineries = getMachineries()
            codes = getCodes()
            intervals = getIntervals()

            if user_input.upper() == "A":
                for _file in srcData["files"]:
                    processDone = generateUJData(
                        _file["excelFile"],
                        vessels,
                        machineries,
                        codes,
                        intervals,
                        debugMode,
                        _file["keys"],
                        _file["type"],
                    )
            elif user_input.upper() == "D":
                for _file in srcData["files"]:
                    if _file["type"] == "deck":
                        processDone = generateUJData(
                            _file["excelFile"],
                            vessels,
                            machineries,
                            codes,
                            intervals,
                            debugMode,
                            _file["keys"],
                            _file["type"],
                        )
            elif user_input.upper() == "E":
                for _file in srcData["files"]:
                    if _file["type"] == "engine":
                        processDone = generateUJData(
                            _file["excelFile"],
                            vessels,
                            machineries,
                            codes,
                            intervals,
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
                processDone = generateUJData(
                    srcData["files"][int(user_input) - 1]["excelFile"],
                    vessels,
                    machineries,
                    codes,
                    intervals,
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


def update_jobs_all(
    srcData: dict,
    vessels: list,
    machineries: list,
    codes: list,
    intervals: list,
    debugMode: bool,
    folder_name: str,
):
    try:
        resetCleanedList()

        console.print("\n\n📝 Update Jobs", highlight=False)

        book = Workbook()
        sheet = book.active
        sheet.append(uj_header)

        for _file in track(
            srcData["files"],
            description="🟢 [bold green]Processing [/bold green]",
        ):
            _ = generateUJData(
                _file["excelFile"],
                vessels,
                machineries,
                codes,
                intervals,
                debugMode,
                _file["keys"],
                _file["type"],
                False,
                False,
                sheet,
            )

        creation_folder = "./res/AIO/" + folder_name + "/update_jobs/"
        if not os.path.exists(creation_folder):
            os.makedirs(creation_folder)
        _filename = "AIO (Update Jobs)" + ".xlsx"
        saveExcelFile(book, _filename, creation_folder)

        global uj_count
        console.print(
            "🔵 Total Encoded Data: " + str(uj_count) + " Row(s)",
            style="bold cyan",
            highlight=False,
        )
        value = getMinVal(uj_count)
        console.print(
            "🟣 Min Rows Per Excel: " + str(value) + " Row(s)",
            style="bold magenta",
            highlight=False,
        )

        excel_count = 1
        global base
        if uj_count >= base:
            excel_count = splitAIO(creation_folder, _filename, "Update Jobs", value)

        console.print(
            "🟡 Total File Created: " + str(excel_count) + " File(s)",
            style="bold yellow",
            highlight=False,
        )
        console.print("📥 Completed", style="info", highlight=False)
    except Exception as e:
        if debugMode:
            logger.exception(e, stack_info=True)
