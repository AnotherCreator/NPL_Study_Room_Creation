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

import logging
import xlsxwriter

WORKBOOK = xlsxwriter.Workbook()
LOGGER = logging.getLogger()

MONTHS = ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]

COL_NAMES = {
        "D", "F", "H", "J", "L", "N", "P", "Q"
}

ROW_NAMES = {
        "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"
}

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

"""
MAIN HEADER FORMATTING PROPERTIES
"""
GENERAL_HEADER = WORKBOOK.add_format(
        {"bold": True, "font": "Calibri", "font_size": 12, "align": "center"}
)

CONFERENCE_ROOM_HEADER = WORKBOOK.add_format(
        {"bold": True, "font": "Calibri", "bg_color": "00B0F0", "font_size": 12, "align": "center"}
)

CAPACITY_TWO = WORKBOOK.add_format(
        {"bold": True, "font": "Calibri", "bg_color": "red", "font_size": 12, "align": "center"}
)

CAPACITY_FIVE = WORKBOOK.add_format(
        {"bold": True, "font": "Calibri", "bg_color": "yellow", "font_size": 12, "align": "center"}
)

CAPACITY_SIX = WORKBOOK.add_format(
        {"bold": True, "font": "Calibri", "bg_color": "lime", "font_size": 12, "align": "center"}
)