#!/usr/bin/env python3

import argparse
import sys
import pandas


def parse_args(args_list: list) -> dict:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--file-path",
        "-f",
        required=True,
        help="The excel file of your data"
    )
    parser.add_argument(
        "--name",
        "-n",
        required=True,
        help="The name to search for"
    )
    parser.add_argument(
        "--column",
        "-c",
        required=True,
        help="The column to search for name"
    )
    parser.add_argument(
        "--sheet",
        "-s",
        required=False,
        default="Sheet1",
        help="Sheet name of data file"
    )
    parser.add_argument(
        "--result-column",
        "-rc",
        required=False,
        help="The column of the result to be printed"
    )
    return parser.parse_args(args_list)


def main(options: dict):
    sheet_df = pandas.read_excel(options.file_path, sheet_name=options.sheet)
    matching_rows = sheet_df[options.column].str.contains(options.name.lower())
    if options.result_column:
        print(sheet_df[matching_rows][options.result_column].values)
    else:
        print(sheet_df[matching_rows])


if __name__ == "__main__":
    opts = parse_args(sys.argv[1:])
    main(opts)
