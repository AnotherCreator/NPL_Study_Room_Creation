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
    THIS FILE WILL CONTAIN MOST OF THE FORMATTING FUNCTIONS
"""
from src.constants import \
    ROOM_LABELS, COL_NAMES, ROW_NAMES, \
    WEEKDAY_HOURS, SUN_SCHOOL_HOURS, SUN_SUMMER_HOURS


def get_general_header(wb):
    return wb.add_format(
        {"bold": True, "font": "Calibri", "font_size": 12, "align": "center"}
    )


def get_conference_room_header(wb):
    return wb.add_format(
        {"bold": True, "font": "Calibri", "bg_color": "00B0F0", "font_size": 12, "align": "center", }
    )


def get_capacity_two_header(wb):
    return wb.add_format(
        {"bold": True, "font": "Calibri", "bg_color": "red", "font_size": 12, "align": "center"}
    )


def get_capacity_five_header(wb):
    return wb.add_format(
        {"bold": True, "font": "Calibri", "bg_color": "yellow", "font_size": 12, "align": "center"}
    )


def get_capacity_six_header(wb):
    return wb.add_format(
        {"bold": True, "font": "Calibri", "bg_color": "lime", "font_size": 12, "align": "center"}
    )


def create_worksheet_headers(wb, ws):
    # Freeze Panes
    ws.freeze_panes("C3")  # This will freeze the study room and time information (Rows 1-2 / Columns A-B)

    # Adjust column widths
    ws.set_column(2, 13, 16.5)  # Study room columns "C:N" with width of 16.5
    ws.set_column(14, 14, 21)  # Conference room column with width of 21
    ws.set_column(15, 16, 16.5)  # SRS / GSR column with width of 16.5

    # Set row 1 column headers
    # Create "Time" header
    ws.merge_range("A1:B2", "Time", get_general_header(wb))

    # Create headers using "study_rooms" function
    for key in ROOM_LABELS:
        ws.write(key + "1", ROOM_LABELS.get(key), get_general_header(wb))

        if ROOM_LABELS.get(key) == "Study Room 5":
            ws.write(key + "2", "Max Capacity: 5", get_capacity_five_header(wb))

        elif ROOM_LABELS.get(key) == "Study Room 9":
            ws.write(key + "2", "Max Capacity: 6", get_capacity_six_header(wb))

        elif ROOM_LABELS.get(key) == "Study Room 10" or ROOM_LABELS.get(key) == "Study Room 11":
            ws.write(key + "2", "Max Capacity: 2", get_capacity_two_header(wb))

        elif ROOM_LABELS.get(key) == "Conference Room":
            # Overwriting general header formatting to include blue bg
            ws.write(key + "1", ROOM_LABELS.get(key), get_conference_room_header(wb))
            ws.write(key + "2", "Max Capacity: 8", get_general_header(wb))

        else:
            ws.write(key + "2", "Max Capacity: 4", get_general_header(wb))
    return


def weekday_interval_times(wb, ws, interval_max=51):
    times = ["9:00", "10:00", "11:00", "12:00", "1:00", "2:00", "3:00", "4:00", "5:00", "6:00", "7:00", "8:00"]
    times_counter = 0

    column_name = "B"  # Letter of column holding 15-min intervals
    interval_align_right = wb.add_format()
    interval_align_right.set_align("right")

    for i in range(3, interval_max, 4):
        min_interval = 0  # 15 min interval

        if min_interval == 45:
            min_interval = 0

        ws.write(column_name + str(i), times[times_counter], interval_align_right)
        min_interval += 15

        for num in range(i + 1, i + 4):
            ws.write("B" + str(num), times[times_counter][:-2] + str(min_interval), interval_align_right)
            min_interval += 15

        times_counter += 1


def sun_reg_interval_times(wb, ws, interval_max=35):
    min_interval = 0  # 15 min interval
    column_name = "B"  # Letter of column holding 15-min intervals

    interval_align_right = wb.add_format()
    interval_align_right.set_align("right")

    for x in range(3, interval_max):
        if min_interval == 60:  # Reset interval counter before each hour
            min_interval = 0

        if x < 7:
            if min_interval == 0:
                ws.write(column_name + str(x), "1:00", interval_align_right)
                min_interval += 15
            else:
                ws.write("B" + str(x), "1:" + str(min_interval), interval_align_right)
                min_interval += 15
        elif 6 < x < 11:
            if min_interval == 0:
                ws.write(column_name + str(x), "2:00", interval_align_right)
                min_interval += 15
            else:
                ws.write("B" + str(x), "2:" + str(min_interval), interval_align_right)
                min_interval += 15
        elif 10 < x < 15:
            if min_interval == 0:
                ws.write(column_name + str(x), "3:00", interval_align_right)
                min_interval += 15
            else:
                ws.write("B" + str(x), "3:" + str(min_interval), interval_align_right)
                min_interval += 15
        elif 14 < x < 19:
            if min_interval == 0:
                ws.write(column_name + str(x), "4:00", interval_align_right)
                min_interval += 15
            else:
                ws.write("B" + str(x), "4:" + str(min_interval), interval_align_right)
                min_interval += 15
        elif 18 < x < 23:
            if min_interval == 0:
                ws.write(column_name + str(x), "5:00", interval_align_right)
                min_interval += 15
            else:
                ws.write("B" + str(x), "5:" + str(min_interval), interval_align_right)
                min_interval += 15
        elif 22 < x < 27:
            if min_interval == 0:
                ws.write(column_name + str(x), "6:00", interval_align_right)
                min_interval += 15
            else:
                ws.write("B" + str(x), "6:" + str(min_interval), interval_align_right)
                min_interval += 15
        elif 26 < x < 31:
            if min_interval == 0:
                ws.write(column_name + str(x), "7:00", interval_align_right)
                min_interval += 15
            else:
                ws.write("B" + str(x), "7:" + str(min_interval), interval_align_right)
                min_interval += 15
        else:
            if min_interval == 0:
                ws.write(column_name + str(x), "8:00", interval_align_right)
                min_interval += 15
            else:
                ws.write("B" + str(x), "8:" + str(min_interval), interval_align_right)
                min_interval += 15


def create_formulas(ws):
    # Formula to count total study room usage
    ws.write("A52", "Users")
    ws.write_formula("B52", "=COUNTA(C3:N50)")

    return


def create_week_day_format(wb, ws):
    # Cell formatting properties
    column_borders = wb.add_format()
    column_borders.set_left(1)
    column_borders.set_right(1)

    column_all_border = wb.add_format()
    column_all_border.set_left(1)
    column_all_border.set_right(1)
    column_all_border.set_bottom(1)

    # Column names that need cell column borders
    for col in COL_NAMES:
        for x in range(3, 51):
            ws.write(col + str(x), "", column_borders)

    # Column names that need cell row borders
    for col in ROW_NAMES:
        n = 6
        for x in range(3, 51):
            if x == n:
                ws.write(col + str(x), "", column_all_border)
                n += 4
            else:
                continue

    for key in WEEKDAY_HOURS:
        ws.merge_range(key, WEEKDAY_HOURS.get(key), get_general_header(wb))

    weekday_interval_times(wb, ws)

    return


def create_sat_format(wb, ws):
    # Cell formatting properties
    column_borders = wb.add_format()
    column_borders.set_left(1)
    column_borders.set_right(1)

    column_all_border = wb.add_format()
    column_all_border.set_left(1)
    column_all_border.set_right(1)
    column_all_border.set_bottom(1)

    # Column names that need cell borders
    for col in COL_NAMES:
        for x in range(3, 35):
            ws.write(col + str(x), "", column_borders)

    # Column names that need cell row borders
    for col in ROW_NAMES:
        n = 6
        for x in range(3, 35):
            if x == n:
                ws.write(col + str(x), "", column_all_border)
                n += 4
            else:
                continue

    # Add hourly cells
    for key in WEEKDAY_HOURS:
        if "A38" in key or "A42" in key or "A46" in key or "A50" in key:
            continue
        ws.merge_range(key, WEEKDAY_HOURS.get(key), get_general_header(wb))

    weekday_interval_times(wb, ws, 35)

    return


def create_sun_format(wb, ws):  # For months excluding June, July, August
    # Cell formatting properties
    column_borders = wb.add_format()
    column_borders.set_left(1)
    column_borders.set_right(1)

    column_all_border = wb.add_format()
    column_all_border.set_left(1)
    column_all_border.set_right(1)
    column_all_border.set_bottom(1)

    # Column names that need cell borders
    for col in COL_NAMES:
        for x in range(3, 35):
            ws.write(col + str(x), "", column_borders)

    # Column names that need cell row borders
    for col in ROW_NAMES:
        n = 6
        for x in range(3, 35):
            if x == n:
                ws.write(col + str(x), "", column_all_border)
                n += 4
            else:
                continue

    for key in SUN_SCHOOL_HOURS:
        ws.merge_range(key, SUN_SCHOOL_HOURS.get(key), get_general_header(wb))

    sun_reg_interval_times(wb, ws)

    return


def create_summer_sun_format(wb, ws):  # For months including June, July, August
    # Cell formatting properties
    column_borders = wb.add_format()
    column_borders.set_left(1)
    column_borders.set_right(1)

    column_all_border = wb.add_format()
    column_all_border.set_left(1)
    column_all_border.set_right(1)
    column_all_border.set_bottom(1)

    # Column names that need cell borders
    for col in COL_NAMES:
        for x in range(3, 19):
            ws.write(col + str(x), "", column_borders)

    # Column names that need cell row borders
    for col in ROW_NAMES:
        n = 6
        for x in range(3, 19):
            if x == n:
                ws.write(col + str(x), "", column_all_border)
                n += 4
            else:
                continue

    for key in SUN_SUMMER_HOURS:
        ws.merge_range(key, SUN_SUMMER_HOURS.get(key), get_general_header(wb))

    sun_reg_interval_times(wb, ws, 19)

    return


def month_total_indirect_formula(ws, n, string_date):
    # Final formula =INDIRECT("'"&A[adjacent cell value being 'n']&"'!"&$B$34)
    indirect_formula_one_half = "=INDIRECT" + '("' + "'" + '"' + "&A"
    indirect_formula_complete_half = '&"' + "'!" + '"&$B$34)'

    # Add worksheet name
    ws.write(n, 0, string_date)

    # Add formula to get total
    ws.write_formula("B" + str(n + 1), indirect_formula_one_half + str(n + 1)
                     + indirect_formula_complete_half)


def create_month_total_format(wb, ws, numeric_date, month):
    # Adjust column width
    ws.set_column(0, 0, 17.5)  # Study room columns "A" with width of 17.5
    ws.set_column(0, 1, 17.5)  # Study room columns "B" with width of 17.5

    # Add header
    ws.write(0, 0, "Worksheet Name", get_general_header(wb))
    ws.write(0, 1, "Totals", get_general_header(wb))

    # Add totals section
    ws.write(32, 0, "Month Totals")

    # Add sum formula
    ws.write_formula("B33", "=SUM(B2:B32)")

    # Add cell where users are stored
    ws.write(33, 0, "Cell Storing Users")
    ws.write(33, 1, "B52")

    #  Format each [Month] Total
    #  Might be able to simplify the following code by removing the month checks since individual month
    #  is already being passed in
    n = 1
    if month == "January":
        for date in numeric_date:
            string_date = date.strftime("%a %b %d")
            if "Jan" in string_date:
                # Add worksheet name
                ws.write(n, 0, string_date)

                # Add indirect formula
                month_total_indirect_formula(ws, n, string_date)

                n += 1
        return
    elif month == "February":
        for date in numeric_date:
            string_date = date.strftime("%a %b %d")
            if "Feb" in string_date:
                ws.write(n, 0, string_date)

                # Add indirect formula
                month_total_indirect_formula(ws, n, string_date)

                n += 1
        return
    elif month == "March":
        for date in numeric_date:
            string_date = date.strftime("%a %b %d")
            if "Mar" in string_date:
                ws.write(n, 0, string_date)

                # Add indirect formula
                month_total_indirect_formula(ws, n, string_date)

                n += 1
        return
    elif month == "April":
        for date in numeric_date:
            string_date = date.strftime("%a %b %d")
            if "Apr" in string_date:
                ws.write(n, 0, string_date)

                # Add indirect formula
                month_total_indirect_formula(ws, n, string_date)

                n += 1
        return
    elif month == "May":
        for date in numeric_date:
            string_date = date.strftime("%a %b %d")
            if "May" in string_date:
                ws.write(n, 0, string_date)

                # Add indirect formula
                month_total_indirect_formula(ws, n, string_date)

                n += 1
        return
    elif month == "June":
        for date in numeric_date:
            string_date = date.strftime("%a %b %d")
            if "Jun" in string_date:
                ws.write(n, 0, string_date)

                # Add indirect formula
                month_total_indirect_formula(ws, n, string_date)

                n += 1
        return
    elif month == "July":
        for date in numeric_date:
            string_date = date.strftime("%a %b %d")
            if "Jul" in string_date:
                ws.write(n, 0, string_date)

                # Add indirect formula
                month_total_indirect_formula(ws, n, string_date)

                n += 1
        return
    elif month == "August":
        for date in numeric_date:
            string_date = date.strftime("%a %b %d")
            if "Aug" in string_date:
                ws.write(n, 0, string_date)

                # Add indirect formula
                month_total_indirect_formula(ws, n, string_date)

                n += 1
        return
    elif month == "September":
        for date in numeric_date:
            string_date = date.strftime("%a %b %d")
            if "Sep" in string_date:
                ws.write(n, 0, string_date)

                # Add indirect formula
                month_total_indirect_formula(ws, n, string_date)

                n += 1
        return
    elif month == "October":
        for date in numeric_date:
            string_date = date.strftime("%a %b %d")
            if "Oct" in string_date:
                ws.write(n, 0, string_date)

                # Add indirect formula
                month_total_indirect_formula(ws, n, string_date)

                n += 1
        return
    elif month == "November":
        for date in numeric_date:
            string_date = date.strftime("%a %b %d")
            if "Nov" in string_date:
                ws.write(n, 0, string_date)

                # Add indirect formula
                month_total_indirect_formula(ws, n, string_date)

                n += 1
        return
    else:
        for date in numeric_date:
            string_date = date.strftime("%a %b %d")
            if "Dec" in string_date:
                ws.write(n, 0, string_date)

                # Add indirect formula
                month_total_indirect_formula(ws, n, string_date)

                n += 1
        return


def create_year_total_format():
    return
