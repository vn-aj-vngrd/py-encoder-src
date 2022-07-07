from app.middleware import *
from app.helpers import *
from app.running_hours_controller import *
from app.sub_categories_controller import *
from app.update_jobs_controller import *


def py_encoder():
    try:
        isError = False
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
            table.add_row("E", "Exit")

            console.print(table)

            if isError:
                console.print("\n:x: Error: " + "You selected an invalid option.")
                file_key = Prompt.ask(
                    ":backhand_index_pointing_right:[yellow blink] Select an option[/yellow blink]"
                )
            else:
                file_key = Prompt.ask(
                    "\n:backhand_index_pointing_right:[yellow blink] Select an option[/yellow blink]"
                )

            isError = False
            if file_key == "R":
                running_hours()

            elif file_key == "S":
                sub_categories()

            elif file_key == "U":
                update_jobs()

            elif file_key == "E":
                break

            else:
                isError = True

    except Exception as e:
        console.print(":x: Error: " + str(e))


py_encoder()
