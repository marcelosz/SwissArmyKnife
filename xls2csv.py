#!/usr/bin/env python
# -*- coding: UTF-8 -*-

"""xls2csv.py - Converts a XLS file to CSV."""

import sys, logging, argparse
import xlrd, csv
import string, textwrap
import unidecode

__author__ = "Marcelo de Souza"
__email__ = "marcelo@marcelosouza.com"
__license__ = "GPL"

#
# globals
# 
logger = ''

#
# main()
#
def main(argv):
    # parse the args
    parser = argparse.ArgumentParser(description="Converts a XLS file to CSV.",
                                     formatter_class=argparse.RawDescriptionHelpFormatter,
                                     epilog=textwrap.dedent("""\
       This script expects the input XLS file to be a well-formed 'CSV-like'
       workbook. The first sheet of the workbook will be converted into a new 
       CSV file. The columns can be choosen through parameters, as well as the
       delimiter. If no specific columns are choosen, all columns are used.

       Example: 
        $ xls2csv.py input.xls output.csv -c \"A,B,Z,AA,AF,AZ\" -d "|"
        - Converts input.xls into ouput.csv, using only data from the XLS columns
          A,B,Z,AA,AF and AZ. The CSV will be delimited by the "|" symbol.

       """))
    parser.add_argument("-c", "--columns", help="List of columns (letters separated by commas) to obtain from the XLS file. If not specified, all columns will be used.")
    parser.add_argument("-d", "--delimiter", help="Delimiter (comma, semi-comma, pipe...).", required=True)
    parser.add_argument("-l", "--loglevel", help="Logging level (DEBUG, INFO or ERROR). If not set, defaults to ERROR.")
    parser.add_argument("infile", help="Input XLS file.", metavar='INFILE')
    parser.add_argument("outfile", help="Output CSV file.", metavar='OUTFILE')
    args = parser.parse_args()
    # configure the logging ("debug") module
    init_logger(args.loglevel)
    # call the converter
    xls2csv(args.infile, args.outfile, args.columns, args.delimiter)

#
# xls2csv() - the dirty work goes here!
#
def xls2csv(infile, outfile, columns_par, delimiter_par):
    sh = ''
    # open the XLS file and get the first sheet only
    try:
        print "> Reading XLS file '{0}'".format(infile)
        sh = xlrd.open_workbook(infile).sheet_by_index(0)
    except:
        logging.error('Ooops... Could not open XLS file')
        sys.exit(2)
    # show info from the first sheet (first is 0)
    print("> Reading sheet '{0}' ({1} rows, {2} columns)".format(sh.name, sh.nrows, sh.ncols))
    # parse the column names / numbers
    n_total_cols = sh.ncols
    n_total_rows = sh.nrows
    n_cols_list = []
    if columns_par == None:
        print "> Columns parameter not specified. Going to obtain data from ALL columns..."
        n_cols_list = range(0,n_total_cols)
    else:
        for c in columns_par.split(','):
            n_cols_list.append(col2num(c))
    # DEBUG - print the column names
    if logger.isEnabledFor(logging.DEBUG):
        logger.debug("List of columns by number:")
        for n in n_cols_list:
            print n, sh.cell_value(rowx=0, colx=n-1) #XLRD columns start at 0
    #
    # set up the CSV header array
    header = []
    header_str = ''
    for n in n_cols_list:
        header.append(sh.cell_value(rowx=0, colx=n-1))
    header_str = ','.join(header)

    print "> CSV header: {0}".format(header_str)
    #
    # iterate and create the CSV file
    print "> Writing CSV file '{0}'".format(outfile)
    with open(outfile, "w") as csv_file:
        csv_row = {}
        if delimiter_par == None:
            delimiter_par = ','
        csv_writer = csv.DictWriter(csv_file, fieldnames=header, delimiter=delimiter_par, quotechar='"', quoting=csv.QUOTE_MINIMAL)
        csv_writer.writeheader()
        for row in range(1, n_total_rows):
            for column in n_cols_list:
                column_name = sh.cell_value(rowx=0, colx=column-1)
                # TODO - handle other cell types. See XLRD docs
                if sh.cell_type(rowx=row, colx=column-1) == xlrd.XL_CELL_NUMBER:
                    value = str(int(sh.cell_value(rowx=row, colx=column-1)))
                    csv_row[column_name] = value  
                else:
                    value = sh.cell_value(rowx=row, colx=column-1)
                    csv_row[column_name] = unidecode.unidecode(value)
            csv_writer.writerow(csv_row)
            csv_row.clear()

#
# col2num() - convert an Excel column name like 'AD' into number '30'
#
def col2num(col):
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num

#
# init_logger() - setup logger (debug messages and stuff...)
#
def init_logger(level):

    global logger
    logger = logging.getLogger()
    if level == 'DEBUG':
        logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
        print "> Debug level: DEBUG"
    elif level == 'INFO':
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
        print "> Debug level: INFO"        
    else:
        logging.basicConfig(level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')


if __name__ == "__main__":
    main(sys.argv[1:])