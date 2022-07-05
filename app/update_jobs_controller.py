from app.helpers import *


def generateUJData(file_name):
    try:
        if not os.path.exists("./data"):
            os.makedirs("./data")

        path = "src/" + file_name
        print("\n📁 File: " + file_name)

        data = pd.read_excel(path, sheet_name=None, index_col=None, header=None)

        xl = pd.ExcelFile(path)
        keys = xl.sheet_names

        book = Workbook()
        sheet = book.active
        sheet.append(uj_header)

        machineries = getMachineries()
        codes = getCodes()
        intervals = getIntervals()
        vessel = str(data[keys[12]].iloc[0, 2])

        for key in keys:
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
                    print("🔃 Processing " + machinery + "...")
                    row = 7

                    # Prepare the sheets
                    # book = Workbook()
                    # sheet = book.active
                    # sheet.append(main_header)

                    while True:

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
                                match = re.match(
                                    r"([a-z]+)([0-9]+)", machinery_code, re.I
                                )
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
                                "sub_categories",
                                file_name,
                                intervals,
                            )

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
                            vessel.strip(),
                            machinery.strip(),
                            str(code).strip(),
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
                    print(
                        '❌ Error: Vessel name or machinery code is missing for sheet "'
                        + key
                        + '"'
                    )

        # create_name = str(file_name[: len(file_name) - 5]).strip()
        # creation_folder = "./res/update_jobs/" + create_name + "/"

        _filename = (
            str(file_name[: len(file_name) - 5]).strip() + " (Update Jobs)" + ".xlsx"
        )
        creation_folder = "./res/update_jobs/"
        if not os.path.exists(creation_folder):
            os.makedirs(creation_folder)
        book.save(creation_folder + _filename)

        print("👌 Done")
    except Exception as e:
        print("❌ Error: " + str(e))


def update_jobs():
    try:
        while True:
            header("👷 Update Jobs")

            files = processSrc("update_jobs")

            file_key = input("\n👉 Select an option: ")

            if file_key != "A":
                file_name = files[int(file_key)]
                generateUJData(file_name)

            else:
                for _file in files:
                    generateUJData(_file)

            if exitApp():
                break

    except Exception as e:
        print("❌ Error: " + str(e))
