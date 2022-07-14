not_included = [
    "Main Menu",
    "Running Hours",
    "MECO Setting",
    "Sheet3",
    "Cylinder Liner Monitoring",
    "ME Exhaust Valve Monitoring",
    "FIVA VALVE Monitoring",
    "Fuel Valve Monitoring",
    "Sheet1",
    "Details",
    " ME Exhaust Valve Monitoring (1",
    " ME Exhaust Valve Monitoring (1)",
    "ME Exhaust Valve Monitoring (1",
    "ME Exhaust Valve Monitoring (1)",
]

sc_header = (
    "vessel",
    "machinery",
    "code",
    "name",
    "description",
    "interval",
    "commissioning_date",
    "last_done_date",
    "last_done_running_hours",
)

uj_header = (
    "vessel",
    "machinery",
    "code",
    "name",
    "description",
    "interval",
    "commissioning_date",
    "last_done_date",
    "last_done_running_hours",
    "instructions",
    "remarks",
    "details",
)

rh_header = ("vessel", "machinery", "running_hours", "updating_date")

interval_header = ("vessel", "machinery", "interval")

debugMode = False

months = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
]

global_cleaned_log_list = []


version_history = [
    "Version 1.0: Initial release (7/11/2022)",
    "Version 1.1: Added functionalities: cleaning the res and log, view version history (7/12/2022)",
    "Version 1.2: Added Error handling: last done hours of sub catergories and update jobs (7/12/2022)",
    "Version 1.3: Fixed: code and date formatting, boolean value in last done running hours (7/12/2022)",
    "Version 1.4: Added functionalities: cleaning src folder (7/12/2022)",
    "Version 1.5: Fixed: ui colors and formatting; Added: encode all modes and new output directory (7/13/2022)",
    "Version 1.6: Fixed: BWMS, BWT, BWMS code (7/13/2022)",
    "Version 1.7: Fixed: simplify encode all ui (7/13/2022)",
    "Version 1.8: Added functionality: encode all modes in one excel file (AIO) (7/14/2022)",
]

global_AIO = False

global_rh_count = 0

global_sc_count = 0

global_uj_count = 0
