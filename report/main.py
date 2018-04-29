import argparse

import openpyxl

import config

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
    parser.add_argument("--output", "-o", help="Path to output pdf")
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

def create_report(wb):
    sheet = wb[config.DEFAULT_SHEET_NAME]
    date_col = sheet[config.DATE_COLUMN]

    for dataset in config.DATASETS:
        title = dataset['title']

        vt_col = sheet[dataset['VT column']]
        mt_col = sheet[dataset['MT column']]

        vt_seamless = remove_seams(date_col, vt_col)

        debug("{0} - {1}: vt {2}, mt {3}".format(
            title,
            date_col[config.DATA_START_ROW].value,
            vt_col[config.DATA_START_ROW].value,
            mt_col[config.DATA_START_ROW].value))



def main():
    args = _parse_arguments()

    if args.debug:
        config.DEBUG = True

    wb = openpyxl.load_workbook(args.input)

    create_report(wb)

if __name__ == '__main__':
    main()