import os
import click
import polib
import openpyxl
from openpyxl.styles import Font, Alignment

try:
    from openpyxl.cell import get_column_letter
except ImportError:
    from openpyxl.utils import get_column_letter

# openpyxl versions < 2.5.0b1
try:
    from openpyxl.cell import WriteOnlyCell
except ImportError:
    from openpyxl.writer.dump_worksheet import WriteOnlyCell
from . import ColumnHeaders


class CatalogFile(click.Path):
    def __init__(self):
        super(CatalogFile, self).__init__(exists=True, dir_okay=False, readable=True)

    def convert(self, value, param, ctx):
        if not os.path.exists(value) and ":" in value:
            # The user passed a <locale>:<path> value
            (locale, path) = value.split(":", 1)
            path = os.path.expanduser(path)
            real_path = super(CatalogFile, self).convert(path, param, ctx)
            return (locale, polib.pofile(real_path, encoding="utf-8-sig"))

        real_path = super(CatalogFile, self).convert(value, param, ctx)
        catalog = polib.pofile(real_path, encoding="utf-8-sig")
        locale = catalog.metadata.get("Language")
        if not locale:
            locale = os.path.splitext(os.path.basename(real_path))[0]
        return (locale, catalog)


# Widths are in range [0, 200]
@click.command()
@click.option(
    "-c",
    "--comments",
    multiple=True,
    default=["extracted"],
    type=click.Choice(["translator", "extracted", "reference", "all"]),
    help="Comments to include in the spreadsheet",
)
@click.option(
    "--width-message-context", type=click.IntRange(0, 200), default=25, help="Width of the namespace", show_default=True
)
@click.option(
    "--width-message-id", type=click.IntRange(0, 200), default=100, help="Width of the message id", show_default=True
)
@click.option("-o", "--output", type=click.File("wb"), default="messages.xlsx", help="Output file", show_default=True)
@click.argument("catalogs", metavar="CATALOG", nargs=-1, required=True, type=CatalogFile())
def main(comments, width_message_context, width_message_id, output, catalogs):
    """
    Convert .PO files to an XLSX file.

    po-to-xls tries to guess the locale for PO files by looking at the
    "Language" key in the PO metadata, falling back to the filename. You
    can also specify the locale manually by adding prefixing the filename
    with "<locale>:". For example: "nl:locales/nl/mydomain.po".
    """

    # Has namespace/group name
    has_msgctxt = False
    for (locale, catalog) in catalogs:
        has_msgctxt = has_msgctxt or any(m.msgctxt for m in catalog)

    # Fonts used
    regular_font = Font(name="Verdana", size=11)
    regular_font_bold = Font(name="Verdana", size=11, bold=True)
    fuzzy_font = Font(italic=True, bold=True)

    alignment_wrap_text = Alignment(wrap_text=True)

    messages = []
    seen = set()
    for (_, catalog) in catalogs:
        for msg in catalog:
            if not msg.msgid or msg.obsolete:
                continue
            if (msg.msgid, msg.msgctxt) not in seen:
                messages.append((msg.msgid, msg.msgctxt, msg))
                seen.add((msg.msgid, msg.msgctxt))

    # NOTE: using optimized mode
    book = openpyxl.Workbook(write_only=True)
    sheet = book.create_sheet(title="Translations")

    def get_cell(value):
        cell = WriteOnlyCell(sheet, value=value)
        cell.font = regular_font
        return cell

    def get_cell_wrapped(value):
        cell = get_cell(value)
        cell.alignment = alignment_wrap_text
        return cell

    def get_cell_bold(value):
        cell = get_cell(value)
        cell.font = regular_font_bold
        return cell

    # NOTE: Because we are using optimized mode we must set these before writing anything
    # https://openpyxl.readthedocs.io/en/stable/optimized.html

    # Reference: https://automatetheboringstuff.com/chapter12/
    # Set size
    sheet.column_dimensions[get_column_letter(1)].width = width_message_context
    sheet.column_dimensions[get_column_letter(2)].width = width_message_id

    # Freeze the first row
    sheet.freeze_panes = "A2"

    # Freeze the first 2 columns
    sheet.freeze_panes = "C2"

    # Write columns header
    row = []
    has_msgctxt_column = has_occurrences_column = has_comment_column = has_tcomment_column = False
    if has_msgctxt:
        has_msgctxt_column = True
        row.append(get_cell_bold(ColumnHeaders.msgctxt))
    row.append(get_cell_bold(ColumnHeaders.msgid))

    # Headers
    if "reference" in comments or "all" in comments:
        has_occurrences_column = True
        row.append(get_cell_bold(ColumnHeaders.occurrences))
    if "extracted" in comments or "all" in comments:
        has_comment_column = True
        row.append(get_cell_bold(ColumnHeaders.comment))
    if "translator" in comments or "all" in comments:
        has_tcomment_column = True
        row.append(get_cell_bold(ColumnHeaders.tcomment))

    # The languages header
    for _, cat in enumerate(catalogs):
        row.append(get_cell_bold(cat[0]))

    # Set fonts
    for i in range(len(row) + 1):
        # index is 1 based
        sheet.column_dimensions[get_column_letter(i + 1)].font = regular_font

    # First row
    sheet.append(row)

    ref_catalog = catalogs[0][1]

    # The rest of the rows
    with click.progressbar(messages, label="Writing catalog to sheet") as todo:
        for (msgid, msgctxt, message) in todo:
            row = []

            # Message namespace
            if has_msgctxt_column:
                row.append(get_cell(msgctxt))

            # Message id
            row.append(get_cell(msgid))

            msg = ref_catalog.find(msgid, msgctxt=msgctxt)

            # Metadata metadata columns
            if has_occurrences_column:
                o = []
                if msg is not None:
                    for (entry, lineno) in msg.occurrences:
                        if lineno:
                            o.append("%s:%s" % (entry, lineno))
                        else:
                            o.append(entry)
                row.append(get_cell(", ".join(o) if o else None))
            if has_comment_column:
                row.append(get_cell(msg.comment if msg is not None else None))
            if has_tcomment_column:
                row.append(get_cell(msg.tcomment if msg is not None else None))

            for cat in catalogs:
                cat = cat[1]
                msg = cat.find(msgid, msgctxt=msgctxt)
                if msg is None:
                    row.append(get_cell(None))
                elif "fuzzy" in msg.flags:
                    cell = WriteOnlyCell(sheet, value=msg.msgstr)
                    cell.font = fuzzy_font
                    row.append(cell)
                else:
                    row.append(get_cell_wrapped(msg.msgstr))

            sheet.append(row)

    book.save(output)


if __name__ == "__main__":
    main()
