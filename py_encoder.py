from app.middleware import *
from app.utils import *
from app.running_hours_controller import *
from app.sub_categories_controller import *
from app.update_jobs_controller import *


def py_encoder():
    try:
        global debugMode
        isError = isExceptionError = False

        if not os.path.exists("./data"):
            os.makedirs("./data")

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

            if debugMode:
                console.print("🛠️ Debug Mode: On", style="success")

            file_key = Prompt.ask(
                ":backhand_index_pointing_right:[yellow blink] Select an option[/yellow blink]"
            )

            isError = False
            if file_key.upper() == "R":
                running_hours(debugMode)
            elif file_key.upper() == "S":
                sub_categories(debugMode)
            elif file_key.upper() == "U":
                update_jobs(debugMode)
            elif file_key.upper() == "D":
                debugMode = debugging()
            elif file_key.upper() == "X":
                showExitCredits()
            else:
                isError = True

    except Exception as e:
        isExceptionError = True
        exceptionMsg = str(e)


py_encoder()
