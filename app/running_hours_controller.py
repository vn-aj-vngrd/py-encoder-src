from app.helpers import *


def generateRHData(file_name: str, machineries: list):
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
        sheet.append(rh_header)

        vessel = str(data[keys[12]].iloc[0, 2])

        for key in keys:
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
                    print("🔃 Processing " + machinery + "...")

                    valid = True

                    running_hours = data[key].iloc[3, 5]
                    if isEmpty(running_hours):
                        valid = False
                        # running_hours = ""

                    updating_date = data[key].iloc[4, 5]
                    if isEmpty(updating_date):
                        valid = False
                        # updating_date = ""

                    if isinstance(updating_date, datetime):
                        updating_date = updating_date.strftime("%d-%b-%y")

                    if valid:
                        rowData = (
                            vessel,
                            machinery,
                            str(running_hours).strip(),
                            str(updating_date).strip(),
                        )
                        sheet.append(rowData)
                else:
                    print("❌ Error: Vessel name or machinery code is missing.")

        # create_name = str(file_name[: len(file_name) - 5]).strip()
        # creation_folder = "./res/running_hours/" + create_name + "/"

        _filename = (
            str(file_name[: len(file_name) - 5]).strip() + " (Running Hours)" + ".xlsx"
        )
        creation_folder = "./res/running_hours/"
        if not os.path.exists(creation_folder):
            os.makedirs(creation_folder)
        book.save(creation_folder + _filename)

        print("👌 Done")
    except Exception as e:
        print("❌ Error: " + str(e))


def running_hours():
    try:
        while True:
            header("⏳ Running Hours")

            files = processSrc("running_hours")

            file_key = input("\n👉 Select an option: ")

            machineries = getMachineries()

            if file_key != "A":
                file_name = files[int(file_key)]
                generateRHData(file_name, machineries)
            else:
                for _file in files:
                    generateRHData(_file, machineries)

            if exitApp():
                break

    except Exception as e:
        print("❌ Error: " + str(e))
