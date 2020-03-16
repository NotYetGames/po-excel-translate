# Translating via spreadsheets

Not all translators are comfortable with using PO-editors such as [Poedit](http://www.poedit.net/) or translation tools like [Transifex](http://trac.transifex.org/). For them this package provides simple tools to
convert PO-files to `xlsx`-files and back again. This also has another benefit:
it is possible to include multiple languages in a single spreadsheet, which can be
helpful when translating to multiple similar languages at the same time (for
example simplified and traditional chinese).

The format for spreadsheets is simple:

* If any message use a message context the first column will specify the
  context. If message contexts are not used this column will be skipped.
* The next (or first) column contains the message id. This is generally the
  canonical text.
* A set of columns for any requested comment types (message occurrences, source
  comments or translator comments).
* A column with the translated text for each locale. Fuzzy translations are
  marked in italic.

**IMPORTANT:** The first row contains the column headers. *``xls2po`` uses these to locale
information in the file, so make sure never to change these!*

# Comparison

NOTE: Original code of this was taken from https://github.com/wichert/po-xls

Advantages of this implementation:
- sane defaults
- the first row and first columns are freezed so that you can always see the source string you want to translate
- customizable options like width/wrap/protected ranges/fonts
- can use the exporter/importer from another python project, you just import the library after installing it:
```py
from pathlib import Path
import po_excel_translate as poet

# po2xls
po_files_to_convert = [
	poet.PortableObjectFile("ro-example.po")
]

poet.PortableObjectFileToXLSX(
	po_files=po_files_to_convert,
	comment_types=[poet.CommentType.SOURCE],
	output_file_path=Path("ro-example.xlsx")
)

# xls2po
poet.XLSXToPortableObjectFile(
	locale="ro",
	input_file_path=Path("ro-example.xlsx"),
	output_file_path=Path("ro-example.po")
)
```

# Install

## From repository
```sh
pip install .
```

## From pypi
```sh
pip install po-excel-translate
```

# Portable Object (.po) to spreadshseet (.xlsx)

Converting one or more PO-files to an xls file is done with the `po2xls`
command:
```sh
po2xls nl.po
```

This will create a new file `messages.xlsx` with the Dutch translations. Multiple
PO files can be specified:
```sh
po2xls -o texts.xlsx zh_CN.po zh_TW.po nl.po
```

This will generate a ``texts.xlsx`` file with all simplified Chinese,
traditional Chinese and Dutch translations.

`po2xls` will guess the locale for a PO file by looking at the `Language`
key in the file metadata, falling back to the filename if no language information
is specified. You can override this by explicitly specifying the locale on the command line. For example:
```sh
po2xls nl:locales/nl/LC_MESSAGES/mydomain.po
```

This will read ``locales/nl/LC_MESSAGES/mydomain.po`` and treat it as Dutch
(``nl`` locale).

You can also use the ``-c`` or ``--comments`` option with one of those choices:
``translator``, ``extracted``, ``reference``, ``all`` to add more column in the
output.

# Spreadshseet (.xlsx) to Portable Object (.po)

Translations can be converted back from a spreadsheet into a PO-file using the `xls2po` command:
```sh
xls2po nl texts.xlsx nl.po
```

This will take the Dutch (`nl`) translations from `texts.xls`, and (re)create an `nl.po` file using those. You can merge those into an existing po-file using a tool like gettext's `msgmerge`.
