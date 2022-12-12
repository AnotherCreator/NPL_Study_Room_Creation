import os.path
from datetime import datetime

import openpyxl
from UliEngineering.Utils.Date import *

from openpyxl import *
from openpyxl.styles import *
import pandas as pd
import xlsxwriter

from os.path import exists
import logging

# GLOBAL VARIABLES
glob_study_rooms = {
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
    "O": "Conference Room"
}


def create_excel_workbook(numeric_date):
    # Check if file was already created
    logging.info("Attempting to create: Test " + str(datetime.now().year + 1) + " Study Room Log.xlsx")
    if exists("Test " + str(datetime.now().year + 1) + " Study Room Log.xlsx"):
        logging.info("Existing file found, file not created")
        logging.info("Attempting to fetch file: "
                     + os.path.basename("Test " + str(datetime.now().year + 1) + " Study Room Log.xlsx"))

        existing_file = os.path.basename("Test " + str(datetime.now().year + 1) + " Study Room Log.xlsx")
        wb = load_workbook(existing_file)

        logging.info("File successfully fetched")
        logging.info("Leaving function: create_excel_workbook()")
        return wb
    # Create Excel file if it does not exist
    else:
        logging.info("Test " + str(datetime.now().year + 1) + " Study Room Log.xlsx NOT FOUND")
        logging.info("Creating: Test " + str(datetime.now().year + 1) + " Study Room Log.xlsx")

        wb = xlsxwriter.Workbook("Test " + str(datetime.now().year + 1) + " Study Room Log.xlsx")

        # List of months
        months = ["January", "February", "March", "April",
                  "May", "June", "July", "August",
                  "September", "October", "November", "December"]

        logging.info("Adding '[DayName] [Month] [DayNumber]' sheets")
        # Add 'DayName Month DayNumber' sheets
        for date in numeric_date:
            string_date = date.strftime("%a %b %d")
            ws = wb.add_worksheet(string_date)
            create_study_rooms(wb, ws)
            create_times(wb, ws)

        logging.info("Adding '[Month] Total' sheets")
        # Add monthly total sheets after
        for month in months:
            worksheet = wb.add_worksheet(month + " Totals")

        logging.info("Adding '[Year] Total' sheet")
        # Add yearly total sheet at the end
        wb.add_worksheet(str(datetime.now().year + 1) + " Totals")

        wb.close()

        logging.info("Leaving function: create_excel_workbook()")
        return wb


def get_days_of_next_year():
    # 'date' format: (year, month, day)
    # Does not include weekday name
    days_in_a_year = []

    for x in all_dates_in_year(datetime.now().year + 1):
        numeric_date = datetime(x.year, x.month, x.day)
        days_in_a_year.append(numeric_date)

    logging.info("Leaving function: get_days_of_next_year()")
    return days_in_a_year


def create_study_rooms(wb, ws):
    # Header formatting properties
    general_headers = wb.add_format({"bold": True})
    general_headers.set_font("Calibri")
    general_headers.set_font_size(14)

    conf_room_headers = wb.add_format({"bold": True})
    conf_room_headers.set_font("Calibri")
    conf_room_headers.set_bg_color("00B0F0")
    conf_room_headers.set_font_size(14)

    capacity_two = wb.add_format({"bold": True})
    capacity_two.set_font("Calibri")
    capacity_two.set_bg_color("red")
    capacity_two.set_font_size(14)

    capacity_five = wb.add_format({"bold": True})
    capacity_five.set_font("Calibri")
    capacity_five.set_bg_color("yellow")
    capacity_five.set_font_size(14)

    capacity_six = wb.add_format({"bold": True})
    capacity_six.set_font("Calibri")
    capacity_six.set_bg_color("lime")
    capacity_six.set_font_size(14)



    # Freeze Panes
    ws.freeze_panes("C3")  # This will freeze the study room and time information (Rows 1-2 / Columns A-B)

    # Adjust column widths
    ws.set_column(2, 13, 17)  # Study room columns "C:N" with width of 17
    ws.set_column(14, 14, 18)  # Conference room column with width of 18
    ws.set_column(15, 16, 10)  # Color legend columns with width of 10

    # Set row 1 column headers
    # Create "Time" header
    ws.write("A1", "Time", general_headers)

    # Create headers using "study_rooms" global var
    for key in glob_study_rooms:
        ws.write(key + "1", glob_study_rooms.get(key), general_headers)
        if glob_study_rooms.get(key) == "Study Room 5":
            ws.write(key + "2", "Max Capacity: 5", capacity_five)
        elif glob_study_rooms.get(key) == "Study Room 9":
            ws.write(key + "2", "Max Capacity: 6", capacity_six)
        elif glob_study_rooms.get(key) == "Study Room 10" or glob_study_rooms.get(key) == "Study Room 11":
            ws.write(key + "2", "Max Capacity: 6", capacity_two)
        elif glob_study_rooms.get(key) == "Conference Room":
            ws.write(key + "1", glob_study_rooms.get(key), conf_room_headers)
            ws.write(key + "2", "Max Capacity: 8", general_headers)
        else:
            ws.write(key + "2", "Max Capacity: 4", general_headers)
    return


def create_times(wb, ws):
    return


def create_cell_borders(wb, ws):
    return


def create_sat_format(wb):
    return


def create_sun_format():
    return


def create_week_day_format():
    return


# Main
if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)
    logging.info("Starting program")

    logging.info("Entering function: get_days_of_next_year()")
    date = get_days_of_next_year()

    logging.info("Entering function: create_excel_workbook()")
    workbook = create_excel_workbook(date)

    logging.info("Ending program")
