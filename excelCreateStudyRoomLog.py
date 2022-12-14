# MIT License
#
# Copyright (c) 2022 Josh R.
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NON-INFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

import logging
import os.path
from datetime import datetime
from os.path import exists

import xlsxwriter
from UliEngineering.Utils.Date import *
from openpyxl import *

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

glob_times_weekdays = {
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

glob_times_sat = {
    "A3:A6": "9:00",
    "A7:A10": "10:00",
    "A11:A14": "11:00",
    "A15:A18": "12:00",
    "A19:A22": "1:00",
    "A23:A26": "2:00",
    "A27:A30": "3:00",
    "A31:A34": "4:00"
}

glob_times_sun_sept_to_may = {
    "A3:A6": "1:00",
    "A7:A10": "2:00",
    "A11:A14": "3:00",
    "A15:A18": "4:00",
    "A19:A22": "5:00",
    "A23:A26": "6:00",
    "A27:A30": "7:00",
    "A31:A34": "8:00"
}

glob_times_sun_june_to_aug = {
    "A3:A6": "1:00",
    "A7:A10": "2:00",
    "A11:A14": "3:00",
    "A15:A18": "4:00"
}


def create_excel_workbook(numeric_date):
    # Check if file was already created
    logging.info("Attempting to create: Test " + str(datetime.now().year) + " Study Room Log.xlsx")
    if exists("Test " + str(datetime.now().year) + " Study Room Log.xlsx"):
        logging.info("Existing file found, file not created")
        logging.info("Attempting to fetch file: "
                     + os.path.basename("Test " + str(datetime.now().year) + " Study Room Log.xlsx"))

        existing_file = os.path.basename("Test " + str(datetime.now().year) + " Study Room Log.xlsx")
        wb = load_workbook(existing_file)

        logging.info("File successfully fetched")
        logging.info("Leaving function: create_excel_workbook()")
        return wb
    # Create Excel file if it does not exist
    else:
        logging.info("Test " + str(datetime.now().year) + " Study Room Log.xlsx NOT FOUND")
        logging.info("Creating: Test " + str(datetime.now().year) + " Study Room Log.xlsx")

        wb = xlsxwriter.Workbook("Test " + str(datetime.now().year) + " Study Room Log.xlsx")

        # List of months
        months = ["January", "February", "March", "April",
                  "May", "June", "July", "August",
                  "September", "October", "November", "December"]

        logging.info("Adding '[DayName] [Month] [DayNumber]' sheets")
        # Add 'DayName Month DayNumber' sheets
        # This will iterate through all known dates of the year
        for date in numeric_date:
            string_date = date.strftime("%a %b %d")
            ws = wb.add_worksheet(string_date)  # Add 'DayName Month DayNumber' sheets

            # Adjust column widths
            ws.set_column(0, 0, 7.5)  # Hourly time width
            ws.set_column(1, 1, 5.5)  # Quarter intervals

            # General worksheet formatting
            create_cell_borders(wb, ws)  # Add cell borders
            create_study_rooms(wb, ws)  # Add room columns and formatting
            create_formulas(wb, ws)  # Add formulas

            # Create y-axis time blocks
            if "Sun" in string_date:  # Create time format for Sundays
                if "Jun" in string_date or "Jul" in string_date or "Aug" in string_date:
                    create_summer_sun_format(wb, ws)
                else:
                    create_sun_format(wb, ws)
            elif "Sat" in string_date:  # Create time format for Saturdays
                create_sat_format(wb, ws)
            else:
                create_week_day_format(wb, ws)  # Create time format for weekdays

        logging.info("Adding '[Month] Total' sheets")
        # Add monthly total sheets after
        for month in months:
            wb.add_worksheet(month + " Totals")

        logging.info("Adding '[Year] Total' sheet")
        # Add yearly total sheet at the end
        wb.add_worksheet(str(datetime.now().year + 1) + " Totals")

        wb.close()

        logging.info("Leaving function: create_excel_workbook()")
        return wb


def get_days_of_current_year(year):
    # 'date' format: (year, month, day)
    # Does not include weekday name
    days_in_a_year = []

    for x in all_dates_in_year(year):
        numeric_date = datetime(x.year, x.month, x.day)
        days_in_a_year.append(numeric_date)

    logging.info("Leaving function: get_days_of_next_year()")
    return days_in_a_year


def create_study_rooms(wb, ws):
    # Header formatting properties
    general_headers = wb.add_format({"bold": True})
    general_headers.set_font("Calibri")
    general_headers.set_font_size(14)
    general_headers.set_align("vcenter")
    general_headers.set_align("center")

    conf_room_headers = wb.add_format({"bold": True})
    conf_room_headers.set_font("Calibri")
    conf_room_headers.set_bg_color("00B0F0")
    conf_room_headers.set_font_size(14)
    general_headers.set_align("vcenter")
    general_headers.set_align("center")

    capacity_two = wb.add_format({"bold": True})
    capacity_two.set_font("Calibri")
    capacity_two.set_bg_color("red")
    capacity_two.set_font_size(14)
    general_headers.set_align("vcenter")
    general_headers.set_align("center")

    capacity_five = wb.add_format({"bold": True})
    capacity_five.set_font("Calibri")
    capacity_five.set_bg_color("yellow")
    capacity_five.set_font_size(14)
    general_headers.set_align("vcenter")
    general_headers.set_align("center")

    capacity_six = wb.add_format({"bold": True})
    capacity_six.set_font("Calibri")
    capacity_six.set_bg_color("lime")
    capacity_six.set_font_size(14)
    general_headers.set_align("vcenter")
    general_headers.set_align("center")

    # Freeze Panes
    ws.freeze_panes("C3")  # This will freeze the study room and time information (Rows 1-2 / Columns A-B)

    # Adjust column widths
    ws.set_column(2, 13, 17)  # Study room columns "C:N" with width of 17
    ws.set_column(14, 14, 18)  # Conference room column with width of 18
    ws.set_column(15, 16, 10)  # Color legend columns with width of 10

    # Set row 1 column headers
    # Create "Time" header
    ws.merge_range("A1:B2", "Time", general_headers)

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


def create_cell_borders(wb, ws):
    # Cell formatting properties
    column_borders = wb.add_format({"bold": True})
    column_borders.set_left(1)
    column_borders.set_right(1)

    return


def create_formulas(wb, ws):
    # Formula to count total study / conference room occupants in a day
    ws.write("P3", "#users")
    ws.write_formula("Q3", "=COUNTA(C3:O50)")

    return


def create_sat_format(wb, ws):
    # Header formatting properties
    general_headers = wb.add_format({"bold": True})
    general_headers.set_font("Calibri")
    general_headers.set_font_size(14)
    general_headers.set_align("vcenter")
    general_headers.set_align("center")

    # Add hourly cells
    for key in glob_times_weekdays:
        if "A38" in key or "A42" in key or "A46" in key or "A50" in key:
            continue
        ws.merge_range(key, glob_times_weekdays.get(key), general_headers)

    return


def create_sun_format(wb, ws):
    # Header formatting properties
    general_headers = wb.add_format({"bold": True})
    general_headers.set_font("Calibri")
    general_headers.set_font_size(14)
    general_headers.set_align("vcenter")
    general_headers.set_align("center")

    for key in glob_times_sun_sept_to_may:
        ws.merge_range(key, glob_times_sun_sept_to_may.get(key), general_headers)

    return


def create_summer_sun_format(wb, ws):
    # Header formatting properties
    general_headers = wb.add_format({"bold": True})
    general_headers.set_font("Calibri")
    general_headers.set_font_size(14)
    general_headers.set_align("vcenter")
    general_headers.set_align("center")

    for key in glob_times_sun_june_to_aug:
        ws.merge_range(key, glob_times_sun_june_to_aug.get(key), general_headers)

    return


def create_week_day_format(wb, ws):
    # Header formatting properties
    general_headers = wb.add_format({"bold": True})
    general_headers.set_font("Calibri")
    general_headers.set_font_size(14)
    general_headers.set_align("vcenter")
    general_headers.set_align("center")

    for key in glob_times_weekdays:
        ws.merge_range(key, glob_times_weekdays.get(key), general_headers)

    ws.write("B3", "9:00")
    ws.write("B4", "9:15")


# Main
if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)
    logging.info("Starting program")

    userinput_year = int(input("\nEnter the year (E.G. 2023) that you wish to create the Study Room Log for: \n"))

    logging.info("Entering function: get_days_of_current_year()")
    date = get_days_of_current_year(userinput_year)

    logging.info("Entering function: create_excel_workbook()")
    workbook = create_excel_workbook(date)

    logging.info("Ending program")
