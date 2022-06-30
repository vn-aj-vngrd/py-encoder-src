from app.helpers import *


def generateIntervalData(file_name):
    try:
        if not os.path.exists("./data"):
            os.makedirs("./data")

        path = "src/" + file_name
        print("Excel File: " + file_name)

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
        intervals.append("")
        print(intervals)

        # Iterate through the sheets
        for key in keys:
            if key not in not_included:
                print(str(key).rstrip())

                # Vessel Name
                vessel = str(data[key].iloc[0, 2])

                machinery: str = getMachinery(
                    str(data[key].iloc[2, 5]).rstrip(), key, "main", file_name
                )

                if not pd.isna(vessel):
                    # Start traversing the data on row 7
                    row = 7
                    is_Valid = True

                    while is_Valid:

                        # Interval
                        interval = str(data[key].iloc[row, 3])

                        if pd.isna(data[key].iloc[row, 0]):
                            is_Valid = False
                            break

                        if pd.isna(data[key].iloc[row, 3]):
                            interval = ""

                        # Check if the interval is hours
                        if not re.search("[a-zA-Z]", interval) and interval != "":
                            interval = interval + " Hours"

                        # If the interval is unique then append
                        if interval.rstrip() not in intervals:
                            intervals.append(interval)
                            rowData = (
                                vessel.rstrip(),
                                machinery.rstrip(),
                                interval.rstrip(),
                            )
                            sheet.append(rowData)

                        row += 1

        create_name = file_name[: len(file_name) - 4]
        creation_folder = "./res/interval/" + create_name
        if not os.path.exists(creation_folder):
            os.makedirs(creation_folder)
        book.save(creation_folder + "/" + file_name)

    except Exception as e:
        print("Error: " + str(e))


def intervalChecker():
    try:
        while True:
            files = processSrc("interval")
            if len(files) == 0:
                break

            file_key = input("\nSelect an option: ")

            if file_key != "A":
                file_name = files[int(file_key)]
                generateIntervalData(file_name)
            else:
                for _file in files:
                    generateIntervalData(_file)

            if exitApp():
                break

    except Exception as e:
        print("Error: " + str(e))
