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
    THIS FILE WILL CONTAIN ALL THE FORMATTING FUNCTIONS
"""

import excel_create_workbook as ecw

def nine_am_interval_time(ws):
    # 9 AM Formatting
    n = 0
    for x in range(3, 7):
        if n == 0:
            ws.write("B" + str(x), "9:00")
            n += 15
        else:
            ws.write("B" + str(x), "9:" + str(n))
            n += 15

def ten_am_interval_time(ws):
    # 9 AM Formatting
    n = 0
    for x in range(7, 11):
        if n == 0:
            ws.write("B" + str(x), "10:00")
            n += 15
        else:
            ws.write("B" + str(x), "10:" + str(n))
            n += 15

# TODO: ADD CELL BORDER FORMATTING
def create_cell_borders(wb, ws):
    # Cell formatting properties
    column_borders = wb.add_format({"bold": True})
    column_borders.set_left(1)
    column_borders.set_right(1)

    return


def create_formulas(wb, ws):
    # Formula to count total study / conference room occupants in a day
    ws.merge_range("A52:B52", "Users")
    ws.write_formula("C52", "=COUNTA(C3:N50)+COUNTA(P3:P50)")

    return


def create_sat_format(wb, ws):
    # Header formatting properties
    general_headers = wb.add_format({"bold": True})
    general_headers.set_font("Calibri")
    general_headers.set_font_size(14)
    general_headers.set_align("vcenter")
    general_headers.set_align("center")

    # Add hourly cells
    for key in ecw.times_weekdays():
        if "A38" in key or "A42" in key or "A46" in key or "A50" in key:
            continue
        ws.merge_range(key, ecw.times_weekdays().get(key), general_headers)

    return


def create_sun_format(wb, ws):
    # Header formatting properties
    general_headers = wb.add_format({"bold": True})
    general_headers.set_font("Calibri")
    general_headers.set_font_size(14)
    general_headers.set_align("vcenter")
    general_headers.set_align("center")

    for key in ecw.times_sun_sept_to_may():
        ws.merge_range(key, ecw.times_sun_sept_to_may().get(key), general_headers)

    return


def create_summer_sun_format(wb, ws):
    # Header formatting properties
    general_headers = wb.add_format({"bold": True})
    general_headers.set_font("Calibri")
    general_headers.set_font_size(14)
    general_headers.set_align("vcenter")
    general_headers.set_align("center")

    for key in ecw.times_sun_june_to_aug():
        ws.merge_range(key, ecw.times_sun_june_to_aug().get(key), general_headers)

    return


def create_week_day_format(wb, ws):
    # Header formatting properties
    general_headers = wb.add_format({"bold": True})
    general_headers.set_font("Calibri")
    general_headers.set_font_size(14)
    general_headers.set_align("vcenter")
    general_headers.set_align("center")

    for key in ecw.times_weekdays():
        ws.merge_range(key, ecw.times_weekdays().get(key), general_headers)

    nine_am_interval_time(ws)
    ten_am_interval_time(ws)