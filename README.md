# mojo-xlsxwriter

`mojo-xlsxwriter` is a standalone XLSX writer with an
[XlsxWriter](https://xlsxwriter.readthedocs.io/)-shaped Python API and a Mojo
kernel for Excel's shared-string XML escaping. It writes ordinary Office Open
XML workbooks without importing XlsxWriter at runtime. XlsxWriter 3.2.9 is a
development dependency and the behavioral reference for the covered subset.

This is a focused port, not a wrapper around upstream and not a complete
reimplementation. The import changes from `xlsxwriter` to
`mojo_xlsxwriter`; covered worksheet methods accept both
zero-based row/column coordinates and XlsxWriter's A1 call forms.

## Coverage

The implemented subset is:

- `Workbook`: path or file-like destinations, context-manager use,
  `add_worksheet`, `add_format`, `get_worksheet_by_name`, `worksheets`,
  `set_properties`, `use_zip64`, and `close`
- `Worksheet`: `write`, `write_string`, `write_number`, `write_boolean`,
  `write_blank`, `write_formula`, `write_datetime`, `write_url`, `write_row`,
  `write_column`, `merge_range`, `set_row`, `set_column`, `freeze_panes`, and
  `autofilter`
- workbook options for the 1900/1904 date systems, default date formats,
  string-to-number/formula/URL dispatch, timezone removal, NaN/Inf formula
  errors, in-memory output, and ZIP64
- cell formats for fonts, solid fills, number formats, alignment, wrapping,
  indentation, rotation, borders, and protection
- insertion-ordered shared strings, duplicate reuse, leading/trailing
  whitespace preservation, Excel `_xHHHH_` control escapes, UTF-8, formulas
  with cached values, external/internal hyperlinks, merged cells, row heights,
  and Excel column-width conversion
- public A1 helpers in `mojo_xlsxwriter.utility`

The tests compare return codes, cell values and types, dates, formulas and
cached values, formatting, dimensions, links, merged ranges, pane/filter
state, and shared-string ordering against real XlsxWriter. The generated
`sharedStrings.xml` is also checked byte-for-byte.

Not covered are charts, images, tables, conditional formatting, data
validation, comments, rich strings, array/dynamic formulas, named ranges,
print/page setup, VBA, encryption, and formula calculation. Excel or another
spreadsheet program calculates formulas when the file opens. XlsxWriter's
`constant_memory` mode uses inline strings rather than shared strings and is
explicitly rejected by this shared-string-focused port.

## Install

From the repository checkout:

```bash
pixi install
pixi run build
pixi run test
```

The build produces `dist/libmojo-xlsxwriter.so`. The Pixi environment sets
`PYTHONPATH=python`, so examples run directly from the checkout.

## Usage

```python
from datetime import date
from mojo_xlsxwriter import Workbook

with Workbook("sales.xlsx") as workbook:
    sheet = workbook.add_worksheet("Sales")
    header = workbook.add_format({
        "bold": True,
        "font_color": "white",
        "bg_color": "#4472C4",
    })
    money = workbook.add_format({"num_format": "$#,##0.00"})
    day = workbook.add_format({"num_format": "yyyy-mm-dd"})

    sheet.write_row("A1", ["Region", "Revenue", "Date"], header)
    sheet.write_row("A2", ["North", 12500.50])
    sheet.write_datetime("C2", date(2026, 7, 30), day)
    sheet.write_number("B2", 12500.50, money)
    sheet.write_formula("B3", "=SUM(B2:B2)", money, 12500.50)
    sheet.freeze_panes(1, 0)
    sheet.autofilter("A1:C2")
```

This produces a standards-based XLSX file whose text cells reference
`xl/sharedStrings.xml`.

## Benchmarks

Measured with `pixi run bench` on an Intel Xeon E5-2697 v4 at 2.30 GHz,
72 logical CPUs, Python 3.13.14, and XlsxWriter 3.2.9. Times are the best of
three warm runs. Ratio is XlsxWriter time divided by mojo-xlsxwriter time.
Every case validates equivalent string IDs/escapes or workbook dimensions,
cell counts, and shared-string counts before printing.

| case | mojo-xlsxwriter | XlsxWriter | ratio | result |
|---|---:|---:|---:|---|
| XML escape (300k unique) | 303.04 ms | 1135.28 ms | 3.75x | faster |
| dedup + escape (600k strings) | 141.07 ms | 214.43 ms | 1.52x | faster |
| XLSX: 150k x 3 strings | 2742.97 ms | 3227.73 ms | 1.18x | faster |
| XLSX: 100k mixed rows | 2055.57 ms | 5559.10 ms | 2.70x | faster |

Python string sequences use one-pass CPython insertion-ordered dictionary
indexing without first encoding every repeated value. UTF-8 offset packing is
prefix-summed in place. Callers that already own packed UTF-8 can use the
low-level Mojo dedup path with contiguous NumPy `uint8` data and `int64`
offsets; both buffers cross the FFI boundary zero-copy. Collision comparisons
and ordinary XML byte runs use unaligned-safe SIMD loads with scalar remainder
handling. XML length measurement and output switch from serial execution to
bounded parallel chunks at 16,384 strings.

No GPU path is included. Shared-string hashing and XML escaping are bytewise,
variable-output operations with well under 2 arithmetic operations per byte;
host/device transfer and compaction overhead would dominate. Workbook and ZIP
serialization also remain CPU-side.

## How it works

Each `write_string` records an occurrence number. At close, a Python
insertion-ordered dictionary assigns the exact shared-string IDs expected by
XlsxWriter. Unique strings are UTF-8 encoded into one contiguous `uint8`
buffer with an `int64` offset table. Python allocates the destination buffer,
then passes 64-bit addresses, element counts, and output capacities through
`ctypes`.

The Mojo kernel reconstructs buffers with
`UnsafePointer[..., AnyOrigin[mut=True]]`, scans every unique string once, and
writes XML entities, Excel control escapes, escaped literal `_xHHHH_`
sequences, and the U+FFFE/U+FFFF noncharacters into the caller-owned output.
No allocation crosses the ABI. Empty inputs stay in Python and output lengths
are checked before decoding.

Python then assembles worksheet, styles, relationships, workbook, document
property, content-type, and shared-string XML parts and writes them with the
standard library's DEFLATE ZIP implementation. Cells are held in sparse
row/column dictionaries and emitted in sorted Excel order, so blank areas do
not consume a dense grid.

MIT.
