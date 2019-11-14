import os
import sys
import time
import click
import polib
import openpyxl
from pathlib import Path
from po_excel_translate import XLSXToPortableObjectFile


@click.command()
@click.argument("locale", required=True)
@click.argument("input_file", type=click.Path(exists=True, readable=True), required=True)
@click.argument("output_file", type=str, required=True)
def main(locale, input_file, output_file):
    """
    Convert a XLS(X) file to a .PO file
    """
    XLSXToPortableObjectFile(locale=locale, input_file_path=Path(str(input_file)), output_file_path=Path(output_file))


if __name__ == "__main__":
    main()
