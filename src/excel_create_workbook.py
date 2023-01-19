# MIT License
#
# Copyright (c) 2022 Josh Reginaldo (https://github.com/AnotherCreator)
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

"""
    THIS FILE WILL HANDLE CREATING THE ACTUAL WORKBOOK AND WORKSHEETS
"""

import logging
import os.path
from datetime import datetime
from os.path import exists

import xlsxwriter  # https://xlsxwriter.readthedocs.io/index.html
from openpyxl import *

import xlsx_formatting

def study_rooms():
    dict_study_rooms = {
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
        "P": "SRS"
    }
    return dict_study_rooms


def times_weekdays():
    dict_times_weekdays = {
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
    return dict_times_weekdays


def times_sat():
    dict_times_sat = {
        "A3:A6": "9:00",
        "A7:A10": "10:00",
        "A11:A14": "11:00",
        "A15:A18": "12:00",
        "A19:A22": "1:00",
        "A23:A26": "2:00",
        "A27:A30": "3:00",
        "A31:A34": "4:00"
    }
    return dict_times_sat


def times_sun_sept_to_may():
    dict_times_sun_sept_to_may = {
        "A3:A6": "1:00",
        "A7:A10": "2:00",
        "A11:A14": "3:00",
        "A15:A18": "4:00",
        "A19:A22": "5:00",
        "A23:A26": "6:00",
        "A27:A30": "7:00",
        "A31:A34": "8:00"
    }
    return dict_times_sun_sept_to_may


def times_sun_june_to_aug():
    dict_times_sun_june_to_aug = {
        "A3:A6": "1:00",
        "A7:A10": "2:00",
        "A11:A14": "3:00",
        "A15:A18": "4:00"
    }
    return dict_times_sun_june_to_aug


def init_workbook(numeric_date, input_year):
    # Check if file was already created
    logging.info("Attempting to create: Test " + str(input_year) + " Study Room Log.xlsx")
    if exists("Test " + str(input_year) + " Study Room Log.xlsx"):
        logging.info("Existing file found, file not created")
        logging.info("Attempting to fetch file: "
                     + os.path.basename("Test " + str(input_year) + " Study Room Log.xlsx"))

        existing_file = os.path.basename("Test " + str(input_year) + " Study Room Log.xlsx")
        wb = load_workbook(existing_file)

        logging.info("File successfully fetched")
        logging.info("Leaving function: create_excel_workbook()")
        return wb
    # Create Excel file if it does not exist
    else:
        logging.info("Test " + str(input_year) + " Study Room Log.xlsx NOT FOUND")
        logging.info("Creating: Test " + str(input_year) + " Study Room Log.xlsx")

        wb = xlsxwriter.Workbook("Test " + str(input_year) + " Study Room Log.xlsx")

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
            xlsx_formatting.create_cell_borders(wb, ws)  # Add cell borders
            create_study_rooms(wb, ws)  # Add room columns and formatting
            xlsx_formatting.create_formulas(wb, ws)  # Add formulas

            # Create y-axis time blocks
            if "Sun" in string_date:  # Create time format for Sundays
                if "Jun" in string_date or "Jul" in string_date or "Aug" in string_date:
                    xlsx_formatting.create_summer_sun_format(wb, ws)
                else:
                    xlsx_formatting.create_sun_format(wb, ws)
            elif "Sat" in string_date:  # Create time format for Saturdays
                xlsx_formatting.create_sat_format(wb, ws)
            else:
                xlsx_formatting.create_week_day_format(wb, ws)  # Create time format for weekdays

        logging.info("Adding '[Month] Total' sheets")
        # Add monthly total sheets
        for month in months:
            wb.add_worksheet(month + " Totals")
            # TODO: ADD FORMATTING / FORMULAS FOR DAILY TOTAL USERS AND MONTH GRAND TOTAL

        logging.info("Adding '[Year] Total' sheet")
        # Add yearly total sheet at the end
        wb.add_worksheet(str(datetime.now().year + 1) + " Totals")
        # TODO: ADD FORMATTING / FORMULAS FOR MONTHLY TOTAL USERS AND YEAR GRAND TOTAL

        wb.close()

        logging.info("Leaving function: create_excel_workbook()")
        return wb


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
    for key in study_rooms():
        ws.write(key + "1", study_rooms().get(key), general_headers)
        if study_rooms().get(key) == "Study Room 5":
            ws.write(key + "2", "Max Capacity: 5", capacity_five)
        elif study_rooms().get(key) == "Study Room 9":
            ws.write(key + "2", "Max Capacity: 6", capacity_six)
        elif study_rooms().get(key) == "Study Room 10" or study_rooms().get(key) == "Study Room 11":
            ws.write(key + "2", "Max Capacity: 6", capacity_two)
        elif study_rooms().get(key) == "Conference Room":
            ws.write(key + "1", study_rooms().get(key), conf_room_headers)
            ws.write(key + "2", "Max Capacity: 8", general_headers)
        elif study_rooms().get(key) == "SRS":
            ws.write(key + "1", study_rooms().get(key), general_headers)
            ws.write(key + "2", "Max Capacity: 4", general_headers)
        else:
            ws.write(key + "2", "Max Capacity: 4", general_headers)

    return
