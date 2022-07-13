from app.utils import *


def generateSCData(
    file_name: str,
    machineries: list,
    codes: list,
    intervals: list,
    debugMode: bool,
    keys: list,
    _type: bool,
    showExtraInfo: bool = True,
):
    try:
        path = "src/" + file_name
        if showExtraInfo:
            console.print("\n\n📑 " + file_name, style="white", highlight=False)

        data = pd.read_excel(path, sheet_name=None, index_col=None, header=None)

        book = Workbook()
        sheet = book.active
        sheet.append(sc_header)

        vessel = str(data[keys[12]].iloc[0, 2])

        error = False

        in_key = keys
        if showExtraInfo:
            in_key = track(keys, description="🟢 [bold green]Processing[/bold green]")

        for key in in_key:
            if key not in not_included:
                machinery_id = str(data[key].iloc[2, 5]).strip()

                machinery = getMachinery(
                    machinery_id,
                    key,
                    "sub_categories",
                    file_name,
                    machineries,
                    vessel,
                )

                if machinery == "Ballast Water Management System" and _type == "engine":
                    machinery = "Ballast Water Treatment System"

                machinery_code = getCode(
                    machinery,
                    key,
                    "sub_categories",
                    file_name,
                    codes,
                    vessel,
                )

                if (
                    not isEmpty(vessel)
                    and not isEmpty(machinery)
                    and not isEmpty(machinery_code)
                ):
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
                                    machinery_code.rstrip() + "-" + col_key[1].lstrip()
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
                                "sub_categories",
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
                                    "sub_categories",
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
                                if str(last_done_date).strip().lower() == "since new":
                                    last_done_date = commissioning_date
                                else:
                                    last_done_date = getFormattedDate(
                                        key,
                                        code,
                                        "sub_categories",
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
                                    "sub_categories",
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

                        #  Insertion
                        rowData = (
                            vessel,
                            machinery,
                            code,
                            str(name).strip(),
                            re.sub("\\s+", " ", str(description.strip())),
                            str(interval).strip(),
                            str(commissioning_date).strip(),
                            str(last_done_date).strip(),
                            str(last_done_running_hours).strip(),
                        )

                        sheet.append(rowData)
                        row += 1

                else:
                    error = True
                    createLog(
                        file_name,
                        vessel,
                        "sub_categories",
                        "❌ Vessel name or machinery code is empty "
                        + "(File: "
                        + file_name
                        + ", Sheet: "
                        + str(key)
                        + ")",
                    )

        creation_folder = "./res/" + vessel + "/sub_categories/"
        if not os.path.exists(creation_folder):
            os.makedirs(creation_folder)

        _filename = (
            str(file_name[: len(file_name) - 5]).strip() + " (Sub Categories)" + ".xlsx"
        )

        saveExcelFile(book, _filename, creation_folder)

        if error and not debugMode and showExtraInfo:
            console.print(
                "❌ Error(s) found, refer to the log folder for more information.",
                style="danger",
                highlight=False,
            )

        if showExtraInfo:
            console.print("📥 Completed", style="info")
        return True

    except Exception as e:
        console.print("❌ " + str(e), style="danger")


def sub_categories(debugMode: bool):
    refresh = True
    processDone = isError = isExceptionError = False
    while True:
        try:
            global cleaned_log_list
            cleaned_log_list.clear()

            if refresh:
                srcData = processSrc("📚 [yellow]Sub Categories[/yellow]", True)
                refresh = False

            header()
            console.print("", srcData["table"], "\n")

            if isExceptionError and debugMode:
                console.print("❌ " + exceptionMsg, style="danger")

            if isError:
                console.print(
                    "❌ " + "You have selected an invalid option.",
                    style="danger",
                )

            if debugMode:
                console.print("🛠️ Debug Mode: On", style="success")

            user_input = Prompt.ask(
                "[blink yellow]👉 Select an option[/blink yellow]",
            )

            machineries = getMachineries()
            codes = getCodes()
            intervals = getIntervals()

            if user_input.upper() == "A":
                for _file in srcData["files"]:
                    processDone = generateSCData(
                        _file["excelFile"],
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
                        processDone = generateSCData(
                            _file["excelFile"],
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
                        processDone = generateSCData(
                            _file["excelFile"],
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
                processDone = generateSCData(
                    srcData["files"][int(user_input) - 1]["excelFile"],
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


def sub_categories_all(
    srcData: dict, machineries: list, codes: list, intervals: list, debugMode: bool
):
    try:
        global cleaned_log_list
        cleaned_log_list.clear()

        console.print("\n\n📚 Sub Categories")

        for _file in track(
            srcData["files"],
            description="🟢 [bold green]Processing [/bold green]",
        ):
            _ = generateSCData(
                _file["excelFile"],
                machineries,
                codes,
                intervals,
                debugMode,
                _file["keys"],
                _file["type"],
                False,
            )

        console.print("📥 Completed", style="info")
    except Exception as e:
        if debugMode:
            logger.exception(e, stack_info=True)
