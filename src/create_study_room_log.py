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
    MAIN DRIVER FILE FOR THE PROGRAM

    Program desc:
    The program will ask the user what year they would like to create the study room log for and create the workbook
    based on the year provided (e.g. "2023")
"""

import logging

from datetime import datetime
from UliEngineering.Utils.Date import all_dates_in_year

import excel_create_workbook as create_workbook


def get_days_of_current_year(year):
    # 'date' format: (year, month, day)
    # Does not include weekday name
    days_in_a_year = []

    for x in all_dates_in_year(year):
        numeric_date = datetime(x.year, x.month, x.day)
        days_in_a_year.append(numeric_date)

    logging.info("Leaving function: get_days_of_next_year()")
    return days_in_a_year


def main():
    logging.basicConfig(level=logging.INFO)
    logging.info("Starting program")

    # Get user input
    userinput_year = int(input("\nEnter the year (E.G. 2023) that you wish to create the Study Room Log for: \n"))

    # Attempt to get a list of all days of 'userinput_year'
    logging.info("Entering file 'excel_create_workbook.py'"
                 "and attempting to call function 'get_days_of_current_year'")
    list_of_dates = get_days_of_current_year(userinput_year)

    # Send the lists of dates to be used for each individual workbook worksheet page
    logging.info("Entering file 'excel_create_workbook.py'"
                 "and attempting to call function 'init_workbook'")
    create_workbook.init_workbook(list_of_dates, userinput_year)

    logging.info("Ending program")


if __name__ == "__main__":
    main()
