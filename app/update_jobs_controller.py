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

        machineries = getMachineries()
        codes = getCodes()
        intervals = getIntervals()

        book = Workbook()
        sheet = book.active
        sheet.append(uj_header)

        for key in keys:
            if key not in not_included:

                vessel = str(data[key].iloc[0, 2])
                machinery_id = str(data[key].iloc[2, 5]).strip()

                machinery_name = getMachinery(
                    machinery_id,
                    key,
                    "update_jobs",
                    file_name,
                    machineries,
                )

                machinery_code = getCode(
                    machinery_name,
                    key,
                    "update_jobs",
                    file_name,
                    codes,
                )

                if (
                    not pd.isna(machinery_name)
                    and (machinery_name != "N/A")
                    and not pd.isna(vessel)
                ):
                    print("🔃 Processing " + machinery_name + "...")
                    row = 7
                    is_Valid = True

                    # Prepare the sheets
                    # book = Workbook()
                    # sheet = book.active
                    # sheet.append(main_header)

                    while is_Valid:

                        rowData = (
                            vessel.strip(),
                            machinery_name.strip(),
                        )

                        for col in range(7):
                            d = data[key].iloc[row, col]

                            if (col == 0) and (
                                (d == "")
                                or (d == " ")
                                or (d == "Note:")
                                or not (has_numbers(str(d)))
                                or (pd.isna(d))
                            ):
                                is_Valid = False
                                break

                            if pd.isna(d):
                                d = ""

                            if col == 0:
                                if machinery_code != "N/A":
                                    if "-" in d:
                                        col_key = d.split("-")
                                        d = (
                                            machinery_code.rstrip()
                                            + "-"
                                            + col_key[1].lstrip()
                                        )
                                    else:
                                        match = re.match(r"([a-z]+)([0-9]+)", d, re.I)
                                        if match:
                                            col_key = match.groups()
                                        d = (
                                            machinery_code.rstrip()
                                            + "-"
                                            + col_key[1].lstrip()
                                        )
                                else:
                                    d = machinery_code

                            if col == 3:
                                if not (re.search("[a-zA-Z]", str(d))) and (d != ""):
                                    d = str(d) + " Hours"

                                machinery_interval = getInterval(
                                    d,
                                    key,
                                    "update_jobs",
                                    file_name,
                                    intervals,
                                )

                                d = machinery_interval

                            if ((col == 4) or (col == 5)) and isinstance(d, datetime):
                                d = d.strftime("%d-%b-%y")

                            d = re.sub("\\s+", " ", str(d))

                            tempTuple = (d.strip(),)
                            rowData += tempTuple

                            if col == 6:
                                instructions = data[key].iloc[row, 10]
                                remarks = data[key].iloc[row, 11]
                                tempTuple = (
                                    instructions,
                                    remarks,
                                )
                                rowData += tempTuple

                        if is_Valid:
                            sheet.append(rowData)
                            row += 1
                        else:
                            break

                    # create_name = str(file_name[: len(file_name) - 4]).strip()
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

        create_name = str(file_name[: len(file_name) - 4]).strip()
        creation_folder = "./res/update_jobs/" + create_name
        if not os.path.exists(creation_folder):
            os.makedirs(creation_folder)
        book.save(creation_folder + "/" + file_name)

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