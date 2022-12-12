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
    headers = wb.add_format({'bold': True})
    headers.set_font('Calibri')
    headers.set_font_size(14)

    # Freeze Panes
    ws.freeze_panes("C3")

    # Adjust column widths
    ws.set_column(2, 13, 17)  # Columns "C:N" with width of 17px

    # Set rows to be manipulated
    ws.write("A1", "Time", headers)
    ws.write("C1", "Study Room 1", headers)
    ws.write("D1", "Study Room 2", headers)

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