import os
import click
import polib
import openpyxl


from pathlib import Path
from . import ColumnHeaders, PortableObjectFile, PortableObjectFileToXLSX


# Widths are in range [0, 200]
@click.command()
@click.option(
    "-c",
    "--comments",
    multiple=True,
    default=["extracted"],
    type=click.Choice(["translator", "extracted", "reference", "all"]),
    help="Comments to include in the spreadsheet",
    show_default=True,
)
@click.option(
    "--width-message-context", type=click.IntRange(0, 200), default=20, help="Width of the namespace", show_default=True
)
@click.option(
    "--width-message-id", type=click.IntRange(0, 200), default=80, help="Width of the message id", show_default=True
)
@click.option("-o", "--output", type=str, default="messages.xlsx", help="Output file", show_default=True)
@click.argument("catalogs_paths", metavar="CATALOG", nargs=-1, required=True, type=click.Path())
def main(comments, width_message_context, width_message_id, output, catalogs_paths):
    """
    Convert .PO files to an XLSX file.

    po-to-xls tries to guess the locale for PO files by looking at the
    "Language" key in the PO metadata, falling back to the filename. You
    can also specify the locale manually by adding prefixing the filename
    with "<locale>:". For example: "nl:locales/nl/mydomain.po".
    """
    po_files = []
    for path in catalogs_paths:
        po_files.append(PortableObjectFile(path))

    output_file_path = Path(output)
    PortableObjectFileToXLSX(
        po_files=po_files,
        comments_type=comments,
        output_file_path=output_file_path,
        width_message_context=width_message_context,
        width_message_id=width_message_id,
    )

    print(f"Generated {output_file_path.absolute()}")


if __name__ == "__main__":
    main()
