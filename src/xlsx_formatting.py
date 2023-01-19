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


def weekday_interval_times(ws, interval_max=51):
    min_interval = 0  # 15 min interval
    column_name = "B"  # Letter of column holding 15-min intervals
    for x in range(3, interval_max):
        if min_interval == 60:  # Reset interval counter before each hour
            min_interval = 0

        if x < 7:
            if min_interval == 0:
                ws.write(column_name + str(x), "9:00")
                min_interval += 15
            else:
                ws.write("B" + str(x), "9:" + str(min_interval))
                min_interval += 15
        elif 6 < x < 11:
            if min_interval == 0:
                ws.write(column_name + str(x), "10:00")
                min_interval += 15
            else:
                ws.write("B" + str(x), "10:" + str(min_interval))
                min_interval += 15
        elif 10 < x < 15:
            if min_interval == 0:
                ws.write(column_name + str(x), "11:00")
                min_interval += 15
            else:
                ws.write("B" + str(x), "11:" + str(min_interval))
                min_interval += 15
        elif 14 < x < 19:
            if min_interval == 0:
                ws.write(column_name + str(x), "12:00")
                min_interval += 15
            else:
                ws.write("B" + str(x), "12:" + str(min_interval))
                min_interval += 15
        elif 18 < x < 23:
            if min_interval == 0:
                ws.write(column_name + str(x), "1:00")
                min_interval += 15
            else:
                ws.write("B" + str(x), "1:" + str(min_interval))
                min_interval += 15
        elif 22 < x < 27:
            if min_interval == 0:
                ws.write(column_name + str(x), "2:00")
                min_interval += 15
            else:
                ws.write("B" + str(x), "2:" + str(min_interval))
                min_interval += 15
        elif 26 < x < 31:
            if min_interval == 0:
                ws.write(column_name + str(x), "3:00")
                min_interval += 15
            else:
                ws.write("B" + str(x), "3:" + str(min_interval))
                min_interval += 15
        elif 30 < x < 35:
            if min_interval == 0:
                ws.write(column_name + str(x), "4:00")
                min_interval += 15
            else:
                ws.write("B" + str(x), "4:" + str(min_interval))
                min_interval += 15
        elif 34 < x < 39:
            if min_interval == 0:
                ws.write(column_name + str(x), "5:00")
                min_interval += 15
            else:
                ws.write("B" + str(x), "5:" + str(min_interval))
                min_interval += 15
        elif 38 < x < 43:
            if min_interval == 0:
                ws.write(column_name + str(x), "6:00")
                min_interval += 15
            else:
                ws.write("B" + str(x), "6:" + str(min_interval))
                min_interval += 15
        elif 42 < x < 47:
            if min_interval == 0:
                ws.write(column_name + str(x), "7:00")
                min_interval += 15
            else:
                ws.write("B" + str(x), "7:" + str(min_interval))
                min_interval += 15
        else:
            if min_interval == 0:
                ws.write(column_name + str(x), "8:00")
                min_interval += 15
            else:
                ws.write("B" + str(x), "8:" + str(min_interval))
                min_interval += 15


def sun_reg_interval_times(ws, interval_max=35):
    min_interval = 0  # 15 min interval
    column_name = "B"  # Letter of column holding 15-min intervals
    for x in range(3, interval_max):
        if min_interval == 60:  # Reset interval counter before each hour
            min_interval = 0

        if x < 7:
            if min_interval == 0:
                ws.write(column_name + str(x), "1:00")
                min_interval += 15
            else:
                ws.write("B" + str(x), "1:" + str(min_interval))
                min_interval += 15
        elif 6 < x < 11:
            if min_interval == 0:
                ws.write(column_name + str(x), "2:00")
                min_interval += 15
            else:
                ws.write("B" + str(x), "2:" + str(min_interval))
                min_interval += 15
        elif 10 < x < 15:
            if min_interval == 0:
                ws.write(column_name + str(x), "3:00")
                min_interval += 15
            else:
                ws.write("B" + str(x), "3:" + str(min_interval))
                min_interval += 15
        elif 14 < x < 19:
            if min_interval == 0:
                ws.write(column_name + str(x), "4:00")
                min_interval += 15
            else:
                ws.write("B" + str(x), "4:" + str(min_interval))
                min_interval += 15
        elif 18 < x < 23:
            if min_interval == 0:
                ws.write(column_name + str(x), "5:00")
                min_interval += 15
            else:
                ws.write("B" + str(x), "5:" + str(min_interval))
                min_interval += 15
        elif 22 < x < 27:
            if min_interval == 0:
                ws.write(column_name + str(x), "6:00")
                min_interval += 15
            else:
                ws.write("B" + str(x), "6:" + str(min_interval))
                min_interval += 15
        elif 26 < x < 31:
            if min_interval == 0:
                ws.write(column_name + str(x), "7:00")
                min_interval += 15
            else:
                ws.write("B" + str(x), "7:" + str(min_interval))
                min_interval += 15
        else:
            if min_interval == 0:
                ws.write(column_name + str(x), "8:00")
                min_interval += 15
            else:
                ws.write("B" + str(x), "8:" + str(min_interval))
                min_interval += 15


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


def create_week_day_format(wb, ws):
    # Header formatting properties
    general_headers = wb.add_format({"bold": True})
    general_headers.set_font("Calibri")
    general_headers.set_font_size(14)
    general_headers.set_align("vcenter")
    general_headers.set_align("center")

    for key in ecw.times_weekdays():
        ws.merge_range(key, ecw.times_weekdays().get(key), general_headers)

    weekday_interval_times(ws)

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

    weekday_interval_times(ws, 35)

    return


def create_sun_format(wb, ws):  # For months excluding June, July, August
    # Header formatting properties
    general_headers = wb.add_format({"bold": True})
    general_headers.set_font("Calibri")
    general_headers.set_font_size(14)
    general_headers.set_align("vcenter")
    general_headers.set_align("center")

    for key in ecw.times_sun_sept_to_may():
        ws.merge_range(key, ecw.times_sun_sept_to_may().get(key), general_headers)

    sun_reg_interval_times(ws)

    return


def create_summer_sun_format(wb, ws):  # For months including June, July, August
    # Header formatting properties
    general_headers = wb.add_format({"bold": True})
    general_headers.set_font("Calibri")
    general_headers.set_font_size(14)
    general_headers.set_align("vcenter")
    general_headers.set_align("center")

    for key in ecw.times_sun_june_to_aug():
        ws.merge_range(key, ecw.times_sun_june_to_aug().get(key), general_headers)

    sun_reg_interval_times(ws, 19)

    return

