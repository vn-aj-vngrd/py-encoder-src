from ensurepip import version


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

cleaned_log_list = []

version = "1.0.0"
