from app.middleware import *
from app.utils import *
from app.running_hours import *
from app.sub_categories import *
from app.update_jobs import *
from app.vessel_machineries import *


def executeAll(debugMode: bool):
    enable_globalAIO()
    # Show header and table
    srcData = processSrc("💯 [yellow]All[/yellow]", False)
    header()
    console.print("", srcData["table"], "\n")

    folder_name = Prompt.ask(
        ":backhand_index_pointing_right:[yellow blink] Enter AIO name[/yellow blink]"
    )

    updateFolderName(folder_name)

    # Get data
    vessels = getVessels()
    machineries = getMachineries()
    codes = getCodes()
    intervals = getIntervals()
    dates = getDates()

    # Execute the modes
    running_hours_all(srcData, vessels, machineries, debugMode, folder_name)
    sub_categories_all(
        srcData, vessels, machineries, codes, intervals, dates, debugMode, folder_name
    )
    update_jobs_all(
        srcData, vessels, machineries, codes, intervals, dates, debugMode, folder_name
    )
    vessel_machineries_all(srcData, vessels, machineries, debugMode, folder_name)

    # Prompt for exit
    promptExit()
    disable_globalAIO()
    resetFolderName()


def py_encoder():
    try:
        global debugMode
        isError = isExceptionError = isClean = isEmpty = False

        if not os.path.exists("./data"):
            os.makedirs("./data")

        if not os.path.exists("./src"):
            os.makedirs("./src")

        while True:
            header()
            mainMenu()

            if isExceptionError and debugMode:
                console.print("❌ " + exceptionMsg, style="danger")

            if isError:
                console.print(
                    "❌ Error: " + "You have selected an invalid option.",
                    style="danger",
                )

            if isClean:
                console.print(
                    "✅ Res and log folder cleaned successfully.", style="success"
                )
                isClean = False

            if isEmpty:
                console.print("✅ Src folder emptied successfully.", style="success")
                isEmpty = False

            if debugMode:
                console.print("🛠️ Debug Mode: On", style="success")

            file_key = Prompt.ask(
                ":backhand_index_pointing_right:[yellow blink] Select an option[/yellow blink]"
            )

            isError = False
            if file_key.upper() == "A":
                executeAll(debugMode)
            elif file_key.upper() == "R":
                running_hours(debugMode)
            elif file_key.upper() == "S":
                sub_categories(debugMode)
            elif file_key.upper() == "U":
                update_jobs(debugMode)
            elif file_key.upper() == "V":
                vessel_machineries(debugMode)
            elif file_key.upper() == "C":
                isClean = cleanResLog()
            elif file_key.upper() == "D":
                debugMode = debugging()
            elif file_key.upper() == "E":
                isEmpty = emptySrc()
            elif file_key.upper() == "H":
                displayVersionHistory()
            elif file_key.upper() == "X":
                showExitCredits()
            else:
                isError = True

    except Exception as e:
        isExceptionError = True
        exceptionMsg = str(e)


py_encoder()
