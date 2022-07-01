from app.helpers import *


def generateRHData(file_name):
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
        sheet.append(rh_header)

        machineries = getMachineries()

        # Iterate through the sheets
        for key in keys:
            if key not in not_included:
                print("🔃 Processing " + str(key).rstrip() + "...")

                # Vessel Name
                vessel = str(data[key].iloc[0, 2])

                # Machinery Name
                machinery = getMachinery(
                    str(data[key].iloc[2, 5]),
                    key,
                    "running_hours",
                    file_name,
                    machineries,
                )

                if (
                    not pd.isna(machinery["name"])
                    and not pd.isna(vessel)
                    and (machinery["name"] != "N/A")
                ):

                    # Running Hours
                    if not pd.isna(data[key].iloc[3, 5]):
                        running_hours = str(data[key].iloc[3, 5])
                    else:
                        running_hours = ""

                    # Updated At
                    if not pd.isna(data[key].iloc[4, 5]):
                        updating_date = data[key].iloc[4, 5].strftime("%d-%b-%y")
                    else:
                        updating_date = " "

                    rowData = (
                        vessel.rstrip(),
                        machinery["name"].rstrip(),
                        running_hours.rstrip(),
                        updating_date.rstrip(),
                    )
                    sheet.append(rowData)
                else:
                    print("❌ Error: Vessel name or machinery code is missing.")

        create_name = str(file_name[: len(file_name) - 4]).rstrip()
        creation_folder = "./res/running_hours/" + create_name
        if not os.path.exists(creation_folder):
            os.makedirs(creation_folder)
        book.save(creation_folder + "/" + create_name + "csv")

        print("👌 Done")
    except Exception as e:
        print("❌ Error: " + str(e))


def runningHours():
    try:
        while True:
            header("⏳ Running Hours")

            files = processSrc("running_hours")

            file_key = input("\n👉 Select an option: ")

            if file_key != "A":
                file_name = files[int(file_key)]
                generateRHData(file_name)
            else:
                for _file in files:
                    generateRHData(_file)

            if exitApp():
                break

    except Exception as e:
        print("❌ Error: " + str(e))
