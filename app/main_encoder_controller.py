from app.helpers import *


def generateMainData(file_name):
    try:
        if not os.path.exists("./data"):
            os.makedirs("./data")

        path = "src/" + file_name
        print("\n📁 File: " + file_name)

        # Read the data
        data = pd.read_excel(path, sheet_name=None, index_col=None, header=None)

        # Get the keys
        xl = pd.ExcelFile(path)
        keys = xl.sheet_names

        interval_names = getIntervals(0)
        interval_ids = getIntervals(1)

        interval_names.append("")
        interval_ids.append("")

        machineries = getMachineries()

        # Iterate through the sheets
        for key in keys:
            if key not in not_included:
                print("🔃 Processing " + str(key).rstrip() + "...")

                # Vessel Name
                vessel = str(data[key].iloc[0, 2])

                # Default Machinery Name: machinery = data[key].iloc[2, 2]
                # Machinery Name using the machinery code
                machinery = getMachinery(
                    str(data[key].iloc[2, 5]).rstrip(),
                    key,
                    "main_encoder",
                    file_name,
                    machineries,
                    vessel,
                )

                if (
                    not pd.isna(machinery)
                    and not pd.isna(vessel)
                    and (machinery != "N/A")
                ):
                    # Start traversing the data on row 7
                    row = 7
                    is_Valid = True

                    # Prepare the sheets
                    book = Workbook()
                    sheet = book.active

                    sheet.append(main_header)

                    while is_Valid:

                        rowData = (
                            vessel.rstrip(),
                            machinery.rstrip(),
                        )

                        for col in range(7):
                            d = data[key].iloc[row, col]

                            if (pd.isna(d)) and (col == 0):
                                is_Valid = False
                                break

                            if pd.isna(d):
                                d = ""

                            if col == 3:
                                if not (re.search("[a-zA-Z]", str(d))) and (d != ""):
                                    d = str(d) + " Hours"
                                track = [
                                    vessel,
                                    machinery
                                ]
                                
                                d = getInterval(d, interval_ids, interval_names, track)

                            if ((col == 4) or (col == 5)) and isinstance(d, datetime):
                                d = d.strftime("%d-%b-%y")
                            else:
                                d = re.sub("\\s+", " ", str(d))

                            tempTuple = (d.rstrip(),)
                            rowData += tempTuple

                        if is_Valid:
                            sheet.append(rowData)
                            row += 1

                    create_name = file_name[: len(file_name) - 4]
                    creation_folder = "./res/main_encoder/" + create_name
                    if not os.path.exists(creation_folder):
                        os.makedirs(creation_folder)
                    book.save(creation_folder + "/" + key + ".xlsx")
                else:
                    print("❌ Error: Vessel name or machinery code is missing.")

        print("👌 Done")
    except Exception as e:
        print("❌ Error: " + str(e))


def mainEncoder():
    try:
        while True:
            header("💻 Main Encoder")

            files = processSrc("main_encoder")

            file_key = input("\n👉 Select an option: ")

            if file_key != "A":
                file_name = files[int(file_key)]
                generateMainData(file_name)

            else:
                for _file in files:
                    generateMainData(_file)

            if exitApp():
                break

    except Exception as e:
        print("❌ Error: " + str(e))
