from app.middleware import *
from app.helpers import *
from app.running_hours_controller import *
from app.sub_categories_controller import *
from app.update_jobs_controller import *


def py_encoder():
    try:
        isError = False

        global debugMode

        while True:
            clear()
            header()

            table = Table(style="magenta")
            table.add_column(
                "[cyan]Option[/cyan]", justify="center", style="cyan", no_wrap=True
            )
            table.add_column(
                "[cyan]Mode[/cyan]", justify="left", style="cyan", no_wrap=True
            )

            table.add_row("R", "Running Hours")
            table.add_row("S", "Sub Categories")
            table.add_row("U", "Update Jobs")
            if debugMode:
                table.add_row("D", "Deactivate Debug Mode")
            else:
                table.add_row("D", "Activate Debug Mode")
            table.add_row("E", "Exit")

            console.print("\n", table, "\n")

            if isError:
                console.print(
                    ":x: Error: " + "You have selected an invalid option.",
                    style="danger",
                )

            if debugMode:
                console.print("🍏 Debug Mode: On", style="success")

            file_key = Prompt.ask(
                ":backhand_index_pointing_right:[yellow blink] Select an option[/yellow blink]"
            )

            isError = False
            if file_key == "R" or file_key == "r":
                running_hours(debugMode)

            elif file_key == "S" or file_key == "s":
                sub_categories(debugMode)

            elif file_key == "U" or file_key == "u":
                update_jobs(debugMode)

            elif file_key == "D" or file_key == "d":
                debugMode = debugging()

            elif file_key == "E" or file_key == "e":
                clear()
                header()

                console.print(
                    "\n\n💻 Source: "
                    + "[url]https://github.com/vn-aj-vngrd/py-encoder[/url]"
                )
                console.print(
                    "💛 Created by: " + "[warning]Van AJ B. Vanguardia[/warning]\n\n"
                )

                for _ in track(
                    range(100), description="[success]Exiting[/success]\n\n"
                ):
                    time.sleep(0.05)

                sys.exit(0)
            else:
                isError = True

    except Exception as e:
        console.print(":x: Error: " + str(e))


py_encoder()
