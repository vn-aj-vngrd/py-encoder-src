from app.utils import *


def generateVMData(
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
            sheet.append(vm_header)
        incharges = getIncharges(path)
        
        # Encoding
        vessel_id = str(data[keys[12]].iloc[0, 2])
        vessel = getVessel(
            vessel_id,
            "vessel_machineries",
            file_name,
            vessels,
        )

        if isEmpty(vessel):
            error = True
            createLog(
                file_name,
                vessel,
                "vessel_machineries",
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
                        "vessel_machineries",
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
                            "vessel_machineries",
                            "❌ Machinery code is undefined "
                            + "(File: "
                            + file_name
                            + ", Sheet: "
                            + str(key)
                            + ")",
                        )
                    else:
                        incharge_rank = getIncharge(
                            machinery,
                            key,
                            "vessel_machineries",
                            file_name,
                            incharges,
                            vessel,
                        )
                        if isEmpty(incharge_rank):
                            incharge_rank = ""

                        # model
                        model = data[key].iloc[3, 2]
                        if isEmpty(model):
                            model = ""

                        # maker
                        maker = data[key].iloc[4, 2]
                        if isEmpty(maker):
                            maker = ""

                        # TODO: installed_date ?

                        rowData = (
                            vessel,
                            machinery,
                            str(incharge_rank).strip(),
                            str(model).strip(),
                            str(maker).strip(),
                            "",
                        )
                        if separateExcel:
                            sheet.append(rowData)
                        else:
                            _sheet.append(rowData)

                        global rh_count
                        rh_count += 1

            if separateExcel:
                creation_folder = "./res/" + vessel + "/vessel_machineries/"
                _filename = (
                    str(file_name[: len(file_name) - 5]).strip()
                    + " (Vessel Machineries)"
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


def vessel_machineries(debugMode: bool):
    refresh = True
    processDone = isError = isExceptionError = False
    while True:
        try:
            resetCleanedList()

            if refresh:
                srcData = processSrc("⚓ [yellow]Vessel Machineries[/yellow]", True)
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
                    processDone = generateVMData(
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
                        processDone = generateVMData(
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
                        processDone = generateVMData(
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
                processDone = generateVMData(
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
