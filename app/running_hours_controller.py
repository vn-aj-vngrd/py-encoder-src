from app.utils import *


def generateRHData(file_name: str, machineries: list, debugMode: bool, keys: list):
    try:
        path = "src/" + file_name
        console.print("\n\n📂 " + file_name, style="warning")

        data = pd.read_excel(path, sheet_name=None, index_col=None, header=None)

        book = Workbook()
        sheet = book.active
        sheet.append(rh_header)

        vessel = str(data[keys[12]].iloc[0, 2])

        warnings_errors = False

        in_key = track(keys, description="🟢 [bold green]Processing[/bold green]")
        if debugMode:
            in_key = keys

        for key in in_key:
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
                    if debugMode:
                        console.print(
                            "🟢 [bold green]Processing: [/bold green]" + machinery
                        )

                    valid = True

                    running_hours = data[key].iloc[3, 5]
                    if isEmpty(running_hours):
                        # valid = False
                        running_hours = "0"

                    updating_date = data[key].iloc[4, 5]
                    if isEmpty(updating_date):
                        # valid = False
                        updating_date = ""

                    if isinstance(updating_date, datetime):
                        updating_date = updating_date.strftime("%d-%b-%y")

                    # if valid:
                    rowData = (
                        vessel,
                        machinery,
                        str(running_hours).strip(),
                        str(updating_date).strip(),
                    )
                    sheet.append(rowData)
                else:
                    warnings_errors = True
                    if debugMode:
                        console.print(
                            '❌ Vessel name or machinery code is empty for sheet "'
                            + key
                            + '"',
                            style="danger",
                        )

        _filename = (
            str(file_name[: len(file_name) - 5]).strip() + " (Running Hours)" + ".xlsx"
        )
        creation_folder = "./res/running_hours/"
        saveExcelFile(book, _filename, creation_folder)

        if warnings_errors and not debugMode:
            console.print(
                "⚠️ Warnings or Errors found, refer to the bin folder.", style="warning"
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
            if refresh:
                srcData = processSrc(
                    "running_hours", "🏃 [yellow]Running Hours[/yellow]"
                )
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
                console.print("sdasdada")
                isError = True

            if processDone:
                isError = processDone = False
                if promptExit():
                    break

        except Exception as e:
            isExceptionError = True
            exceptionMsg = str(e)
