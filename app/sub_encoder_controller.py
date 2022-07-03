from app.helpers import *


def generateSubData(file_name):
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

        machineries = getMachineries()

        for key in keys:
            if key not in not_included:
                vessel = str(data[key].iloc[0, 2])
                machinery_id = str(data[key].iloc[2, 5])

                machinery_name = getMachinery(
                    machinery_id,
                    key,
                    "sub_encoder",
                    file_name,
                    machineries,
                )

                print("🔃 Processing " + machinery_name + "...")

                if (
                    not pd.isna(machinery_name)
                    and not pd.isna(vessel)
                    and (machinery_name != "N/A")
                ):

                    if not pd.isna(data[key].iloc[3, 5]):
                        running_hours = str(data[key].iloc[3, 5])
                    else:
                        running_hours = ""

                    if not pd.isna(data[key].iloc[4, 5]):
                        updating_date = data[key].iloc[4, 5].strftime("%d-%b-%y")
                    else:
                        updating_date = ""

                    rowData = (
                        vessel,
                        machinery_name,
                        running_hours,
                        updating_date,
                    )
                    sheet.append(rowData)
                else:
                    print("❌ Error: Vessel name or machinery code is missing.")

        create_name = str(file_name[: len(file_name) - 4]).rstrip()
        creation_folder = "./res/sub_encoder/" + create_name
        if not os.path.exists(creation_folder):
            os.makedirs(creation_folder)
        book.save(creation_folder + "/" + file_name)

        print("👌 Done")
    except Exception as e:
        print("❌ Error: " + str(e))


def subEncoder():
    try:
        while True:
            header("⏳ Sub Encoder - Running Hours")

            files = processSrc("sub_encoder")

            file_key = input("\n👉 Select an option: ")

            if file_key != "A":
                file_name = files[int(file_key)]
                generateSubData(file_name)
            else:
                for _file in files:
                    generateSubData(_file)

            if exitApp():
                break

    except Exception as e:
        print("❌ Error: " + str(e))
