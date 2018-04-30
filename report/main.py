#coding: utf-8
import argparse
from datetime import datetime, timedelta

import openpyxl
from scipy import interpolate

import config
from pdf import Pdf

"""
Main module that implements the command-line interface.
"""

class ArgParser(argparse.ArgumentParser):
    """
    Argument parser that displays help on error
    """
    def error(self, message):
        self.print_help()
        sys.stderr.write("error: {}\n".format(message))
        sys.exit(2)


def _parse_arguments():
    parser = ArgParser(
        description="Report generator for building energy consumption",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    parser.add_argument("--input", "-i", help="Path to input xlsx")
    parser.add_argument("--output", "-o",
        default="report.pdf",
        help="Path to output pdf")
    parser.add_argument("--debug", "-d",
        action="store_const", const=True,
        help="Enable debug outputs")
    args = parser.parse_args()
    return args

def debug(message):
    if config.DEBUG:
        print(message)

def remove_seams(date_col, data_col):
    """
    Remove the seams that occur when the meter gets
    swapped out for a different one. Look for two different
    entries on the same day, where the first of the pair
    is the last entry from the old meter, and the second
    is the first entry from the new meter.
    """
    timeseries = [(date_col[config.DATA_START_ROW].value, 0)]

    loop_zip = zip(date_col[config.DATA_START_ROW + 1:],
        data_col[config.DATA_START_ROW:-1],
        data_col[config.DATA_START_ROW + 1:])

    for d, e1, e2 in loop_zip:
        if e1.value is None or e2.value is None:
            continue

        d_prev = timeseries[-1][0]
        val_prev = timeseries[-1][1]

        if d_prev == d.value:
            continue

        row_diff = e2.value - e1.value
        row_val = val_prev + row_diff
        debug("  {0}: {1} -> {2} ({3}): {4}".format(
            d.value, e1.value, e2.value, row_diff, row_val))
        timeseries.append((d.value, row_val))

    return timeseries

def date_to_string(date):
    return date.strftime("%d. %m. %Y")

def month_to_string(date):
    return date.strftime("%m.%Y")

def pretty_float(value):
    return "{:,.2f}".format(value).replace(",", " ").replace(
        ".", ",")

def interpolate_month_starts(vt_data, mt_data):
    """
    Returns the interpolated values of the accummulated
    metered data in kWh on each day 1 of each month.
    vt_data and mt_data are both lists of (date, value) tuples.
    The return value is a list of (date, vt, mt) tuples ordered ascended
    by date.
    """

    date_start = vt_data[0][0]
    date_end = vt_data[-1][0]

    new_dates = []
    di = date_start
    if di.day > 1:
        # start with the first day of the next week
        di += timedelta(days=32 - di.day)
        di = datetime(di.year, di.month, 1)

    while di <= date_end:
        new_dates.append(di)
        di += timedelta(days=31)
        di = datetime(di.year, di.month, 1)

    debug("Interpolating from {} to {}".format(di, date_end))
    new_data = []
    for dataset in [vt_data, mt_data]:
        x = [ d[0].timestamp() for d in dataset ]
        y = [ d[1] for d in dataset ]
        f = interpolate.interp1d(x, y)

        new_x = [ d.timestamp() for d in new_dates ]
        new_y = f(new_x)
        new_data.append(new_y)
        debug("Interpolation result:\n{}".format(new_y))

    return list(zip(new_dates, new_data[0], new_data[1]))

def get_energy_diff(data):
    """
    Compute the differences between data points in the
    time series of tuples (date, vt, mt). The result is
    a list of (date, vt_delta, mt_delta) with vt_delta
    and mt_delta being the amount of energy at date since
    the preceding date point.
    """

    diff_data = [ [ d[0] for d in data[1:] ] ]
    for c in range(1, 3):
        diff = [ b[c] - a[c] for a, b in zip(data[:-1], data[1:]) ]
        diff_data.append(diff)

    return list(zip(diff_data[0], diff_data[1], diff_data[2]))

def last_12_entries_table(pdf, title, data_list):
    rows = [
        ("Datum meritve", "visoka tarifa [kWh]", "nizka tarifa [kWh]",
            "skupaj [kWh]", "glede na lani [kWh]")
    ]
    sizes = 30, 35, 35, 35, 35

    # data from this year
    sub_data = data_list[-12:]
    # data from previous year
    sub_data_ly = data_list[-24:-12]

    for (d, vt, mt), (date_ly, vt_ly, mt_ly) in zip(sub_data, sub_data_ly):
        rows.append(
            (date_to_string(d),
                pretty_float(vt),
                pretty_float(mt),
                pretty_float(vt + mt),
                pretty_float((vt + mt) - (vt_ly + mt_ly))))

    pdf.add_paragraph("Zadnjih 12 meritev:")
    pdf.add_table(rows, sizes)

def last_12_entries_line_chart(pdf, title, data_list):
    # data from this year
    sub_data = data_list[-12:]
    # data from previous year
    sub_data_ly = data_list[-24:-12]

    series = ["poraba VT", "poraba MT", "poraba skupaj", "poraba lani"]
    labels = []
    values = [ [], [], [], [] ]
    for (date, vt, mt), (date_ly, vt_ly, mt_ly) in zip(sub_data, sub_data_ly):
        labels.append(month_to_string(date - timedelta(days=1)))
        values[0].append(vt)
        values[1].append(mt)
        values[2].append(vt + mt)
        values[3].append(vt_ly + mt_ly)

    pdf.add_line_chart(170, 120, labels, values, series, minv=0)

def last_12_entries_diffs(pdf, data_a, data_b):
    sub_data_a = data_a[-12:]
    sub_data_b = data_b[-12:]

    rows = [
        ("Datum meritve", "razlika VT [kWh]",
            "razlika MT [kWh]", "razlika skupno [kWh]")
    ]
    sizes = 30, 35, 35, 35

    series = list(rows[0][1:])
    labels = []
    values = [ [], [], [] ]

    for (d_a, vt_a, mt_a), (d_b, vt_b, mt_b) in zip(sub_data_a, sub_data_b):
        vt_diff = vt_b - vt_a
        mt_diff = mt_b - mt_a
        sum_diff = (vt_b + mt_b) - (vt_a + mt_a)
        rows.append(
            (date_to_string(d_a),
                pretty_float(vt_diff),
                pretty_float(mt_diff),
                pretty_float(sum_diff)))

        labels.append(month_to_string(d_a - timedelta(days=1)))
        values[0].append(vt_diff)
        values[1].append(mt_diff)
        values[2].append(sum_diff)

    pdf.add_table(rows, sizes)
    pdf.add_line_chart(170, 120, labels, values, series)

def create_report(wb, pdf):
    sheet = wb[config.DEFAULT_SHEET_NAME]
    date_col = sheet[config.DATE_COLUMN]

    all_data_monthly_energy = []

    for dataset in config.DATASETS:
        title = dataset['title']

        vt_col = sheet[dataset['VT column']]
        mt_col = sheet[dataset['MT column']]

        vt_seamless = remove_seams(date_col, vt_col)
        mt_seamless = remove_seams(date_col, mt_col)

        debug("{0} - {1}: vt {2}, mt {3}".format(
            title,
            date_col[config.DATA_START_ROW].value,
            vt_col[config.DATA_START_ROW].value,
            mt_col[config.DATA_START_ROW].value))

        data_months = interpolate_month_starts(vt_seamless, mt_seamless)
        data_monthly_energy = get_energy_diff(data_months)
        all_data_monthly_energy.append(data_monthly_energy)

        pdf.add_paragraph(title)
        last_12_entries_table(pdf, title, data_monthly_energy)
        last_12_entries_line_chart(pdf, title, data_monthly_energy)
        pdf.new_page()

    pdf.add_paragraph("Razlike med B in A")
    data_a = all_data_monthly_energy[0]
    data_b = all_data_monthly_energy[1]
    last_12_entries_diffs(pdf, data_a, data_b)

def set_header(pdf):
    header_lines = [
        config.PDF_HEADER_TITLE,
        "izdelano {}".format(date_to_string(datetime.today()))
    ]
    pdf.set_header(header_lines)

def main():
    args = _parse_arguments()

    if args.debug:
        config.DEBUG = True

    wb = openpyxl.load_workbook(args.input)
    pdf = Pdf(args.output)

    set_header(pdf)
    create_report(wb, pdf)

    pdf.save()

if __name__ == '__main__':
    main()