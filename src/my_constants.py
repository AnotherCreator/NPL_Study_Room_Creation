import logging

LOGGER = logging.getLogger()
MONTHS = ["January", "February", "March", "April",
                  "May", "June", "July", "August",
                  "September", "October", "November", "December"]
COL_NAMES = {"D", "F", "H", "J", "L", "N", "P", "Q"}
ROW_NAMES = {"C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"}

ROOM_LABELS = {
        # FORMAT = Column Letter: Name of room
        "C": "Study Room 1",
        "D": "Study Room 2",
        "E": "Study Room 3",
        "F": "Study Room 4",
        "G": "Study Room 5",
        "H": "Study Room 6",
        "I": "Study Room 7",
        "J": "Study Room 8",
        "K": "Study Room 9",
        "L": "Study Room 10",
        "M": "Study Room 11",
        "N": "Study Room 12",
        "O": "Conference Room",
        "P": "SRS",
        "Q": "GSR"
        # Add more room columns here
    }

WEEKDAY_HOURS = {
        # FORMAT = Cell range: Hour
        "A3:A6": "9:00",
        "A7:A10": "10:00",
        "A11:A14": "11:00",
        "A15:A18": "12:00",
        "A19:A22": "1:00",
        "A23:A26": "2:00",
        "A27:A30": "3:00",
        "A31:A34": "4:00",
        "A35:A38": "5:00",
        "A39:A42": "6:00",
        "A43:A46": "7:00",
        "A47:A50": "8:00"
    }

SAT_HOURS = {
        "A3:A6": "9:00",
        "A7:A10": "10:00",
        "A11:A14": "11:00",
        "A15:A18": "12:00",
        "A19:A22": "1:00",
        "A23:A26": "2:00",
        "A27:A30": "3:00",
        "A31:A34": "4:00"
    }

SUN_SCHOOL_HOURS = {
        "A3:A6": "1:00",
        "A7:A10": "2:00",
        "A11:A14": "3:00",
        "A15:A18": "4:00",
        "A19:A22": "5:00",
        "A23:A26": "6:00",
        "A27:A30": "7:00",
        "A31:A34": "8:00"
    }

SUN_SUMMER_HOURS = {
        "A3:A6": "1:00",
        "A7:A10": "2:00",
        "A11:A14": "3:00",
        "A15:A18": "4:00"
    }