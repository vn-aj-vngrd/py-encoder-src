from app.helpers import *


def generateIntervalData(file_name):
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

        # Prepare the sheets
        book = Workbook()
        sheet = book.active

        # Append the header
        sheet.append(interval_header)

        # Array of intervals
        intervals = getIntervals(1)
        # intervals.append("")

        machineries = getMachineries()

        # Iterate through the sheets
        for key in keys:
            if key not in not_included:
                print("🔃 Processing " + str(key).rstrip() + "...")

                # Vessel Name
                vessel = str(data[key].iloc[0, 2])

                machinery = getMachinery(
                    str(data[key].iloc[2, 5]).rstrip(),
                    key,
                    "interval_checker",
                    file_name,
                    machineries,
                )

                if (
                    not pd.isna(vessel)
                    and not pd.isna(machinery["name"])
                    and (machinery["name"] != "N/A")
                ):
                    # Start traversing the data on row 7
                    row = 7
                    is_Valid = True

                    while is_Valid:

                        # Interval
                        interval = str(data[key].iloc[row, 3])

                        if pd.isna(data[key].iloc[row, 0]):
                            is_Valid = False
                            break

                        # or isinstance(data[key].iloc[row, 3], datetime)
                        if pd.isna(data[key].iloc[row, 3]):
                            interval = ""

                        # Check if the interval is in hour format
                        if not re.search("[a-zA-Z]", interval) and interval != "":
                            interval = interval + " Hours"

                        # If the interval is unique then append
                        if interval.rstrip() not in intervals:
                            intervals.append(interval)
                            rowData = (
                                vessel.rstrip(),
                                machinery["name"].rstrip(),
                                interval.rstrip(),
                            )
                            sheet.append(rowData)
                            print(
                                "⚠️ Warning: "
                                + interval.rstrip()
                                + " is not a valid interval."
                            )

                        row += 1
                else:
                    print("❌ Error: Vessel name or machinery code is missing.")

        create_name = str(file_name[: len(file_name) - 4]).rstrip()
        creation_folder = "./res/interval_checker/" + create_name
        if not os.path.exists(creation_folder):
            os.makedirs(creation_folder)
        book.save(creation_folder + "/" + file_name)

        print("👌 Done")
    except Exception as e:
        print("❌ Error: " + str(e))


def intervalChecker():
    try:
        while True:
            header("⌚ Interval Checker")

            files = processSrc("interval_checker")

            file_key = input("\n👉 Select an option: ")

            if file_key != "A":
                file_name = files[int(file_key)]
                generateIntervalData(file_name)
            else:
                for _file in files:
                    generateIntervalData(_file)

            if exitApp():
                break

    except Exception as e:
        print("❌ Error: " + str(e))
