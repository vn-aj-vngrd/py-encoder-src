from app.helpers import *


def generateUJData(
    file_name: str, machineries: list, codes: list, intervals: list, debugMode: bool
):
    try:
        if not os.path.exists("./data"):
            os.makedirs("./data")

        path = "src/" + file_name
        console.print("\n\n📝 " + file_name)

        data = pd.read_excel(path, sheet_name=None, index_col=None, header=None)

        xl = pd.ExcelFile(path)
        keys = xl.sheet_names

        book = Workbook()
        sheet = book.active
        sheet.append(uj_header)

        vessel = str(data[keys[12]].iloc[0, 2])

        warnings_errors = False

        in_key = track(keys, description="🟢 [success]Processing")
        if debugMode:
            in_key = keys

        for key in in_key:
            if key not in not_included:
                machinery_id = str(data[key].iloc[2, 5]).strip()

                machinery = getMachinery(
                    machinery_id,
                    key,
                    "update_jobs",
                    file_name,
                    machineries,
                )

                machinery_code = getCode(
                    machinery,
                    key,
                    "update_jobs",
                    file_name,
                    codes,
                )

                if (
                    not isEmpty(vessel)
                    and not isEmpty(machinery)
                    and not isEmpty(machinery_code)
                ):
                    if debugMode:
                        console.print("🟢 [success]Processing: [/success]" + machinery)
                    row = 7

                    # Prepare the sheets
                    # book = Workbook()
                    # sheet = book.active
                    # sheet.append(main_header)

                    while True:

                        code = data[key].iloc[row, 0]
                        if not isValid(code):
                            # warnings_errors = True
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

                        name = data[key].iloc[row, 1]
                        if isEmpty(name):
                            name = ""
                        # Manual Override (--Force Fix)
                        if str(code) == "RE-009":
                            name = "EPIRB"

                        description = data[key].iloc[row, 2]
                        if isEmpty(description):
                            description = ""

                        interval = data[key].iloc[row, 3]
                        if isEmpty(interval):
                            interval = ""
                        else:
                            if not (re.search("[a-zA-Z]", str(interval))):
                                interval = str(interval) + " Hours"

                            interval = getInterval(
                                interval,
                                key,
                                "update_jobs",
                                file_name,
                                intervals,
                            )

                            if isEmpty(interval):
                                warnings_errors = True

                        commissioning_date = data[key].iloc[row, 4]
                        if isEmpty(commissioning_date):
                            commissioning_date = ""
                        else:
                            if isinstance(commissioning_date, datetime):
                                commissioning_date = commissioning_date.strftime(
                                    "%d-%b-%y"
                                )

                        last_done_date = data[key].iloc[row, 5]
                        if isEmpty(last_done_date):
                            last_done_date = ""
                        else:
                            if isinstance(last_done_date, datetime):
                                last_done_date = last_done_date.strftime("%d-%b-%y")

                        last_done_running_hours = data[key].iloc[row, 6]
                        if isEmpty(last_done_running_hours):
                            last_done_running_hours = ""

                        instructions = data[key].iloc[row, 10]
                        if isEmpty(instructions):
                            instructions = ""

                        remarks = data[key].iloc[row, 11]
                        if isEmpty(remarks):
                            remarks = ""

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

                        sheet.append(rowData)
                        row += 1

                    # create_name = str(file_name[: len(file_name) - 5]).strip()
                    # creation_folder = "./res/update_jobs/" + create_name
                    # if not os.path.exists(creation_folder):
                    #     os.makedirs(creation_folder)
                    # name_key = str(key).strip()
                    # book.save(creation_folder + "/" + name_key + ".xlsx")

                else:
                    warnings_errors = True
                    if debugMode:
                        console.print(
                            '❌ Vessel name or machinery code is missing for sheet "'
                            + key
                            + '"',
                            style="danger",
                        )

        _filename = (
            str(file_name[: len(file_name) - 5]).strip() + " (Update Jobs)" + ".xlsx"
        )
        creation_folder = "./res/update_jobs/"
        if not os.path.exists(creation_folder):
            os.makedirs(creation_folder)
        book.save(creation_folder + _filename)

        if warnings_errors and not debugMode:
            console.print(
                "⚠️ Warnings or Errors found, refer to the bin folder.", style="warning"
            )

        console.print("📥 Done")
    except Exception as e:
        console.print("❌ " + str(e), style="danger")


def update_jobs(debugMode: bool):
    processDone = isError = False
    while True:
        try:
            clear()
            header()

            files = processSrc(
                "sub_categories",
                "🛠️ [yellow]Update Jobs[/yellow]",
            )

            files_count = len(files)

            if isError:
                console.print(
                    "❌ " + "You have selected an invalid option.",
                    style="danger",
                )

            if debugMode:
                console.print("🍏 Debug Mode: On", style="success")

            file_key = Prompt.ask(
                "[blink yellow]👉 Select an option[/blink yellow]",
            )

            machineries = getMachineries()
            codes = getCodes()
            intervals = getIntervals()

            isError: False
            if file_key == "A" or file_key == "a":
                for _file in files:
                    generateUJData(_file, machineries, codes, intervals, debugMode)
                processDone = True
                isError = False
            elif file_key == "G" or file_key == "g":
                break
            elif int(file_key) >= 1 and int(file_key) <= files_count:
                file_name = files[int(file_key) - 1]
                generateUJData(file_name, machineries, codes, intervals, debugMode)
                processDone = True
                isError = False
            else:
                isError = True

            if processDone and promptExit():
                isError = processDone = False
                break
        except Exception as e:
            isError = True
            if debugMode:
                console.print("❌ " + str(e), style="danger")
