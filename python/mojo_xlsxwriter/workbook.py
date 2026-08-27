"""Standalone XLSX workbook writer with Mojo-backed shared strings."""

from __future__ import annotations

import datetime as dt
import html
import io
import math
import os
import re
import zipfile
from pathlib import Path

from ._lib import deduplicate, escape_xml
from .exceptions import (
    DuplicateWorksheetName,
    FileCreateError,
    InvalidWorksheetName,
)
from .format import Format
from .utility import xl_range, xl_range_abs, xl_rowcol_to_cell
from .worksheet import Cell, Worksheet

XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

_COLOR_NAMES = {
    "black": "000000",
    "blue": "0000FF",
    "brown": "800000",
    "cyan": "00FFFF",
    "gray": "808080",
    "green": "008000",
    "lime": "00FF00",
    "magenta": "FF00FF",
    "navy": "000080",
    "orange": "FF6600",
    "pink": "FF00FF",
    "purple": "800080",
    "red": "FF0000",
    "silver": "C0C0C0",
    "white": "FFFFFF",
    "yellow": "FFFF00",
}
_BORDER_STYLES = {
    1: "thin",
    2: "medium",
    3: "dashed",
    4: "dotted",
    5: "thick",
    6: "double",
    7: "hair",
    8: "mediumDashed",
    9: "dashDot",
    10: "mediumDashDot",
    11: "dashDotDot",
    12: "mediumDashDotDot",
    13: "slantDashDot",
}


def _attr(value) -> str:
    return html.escape(str(value), quote=True)


def _tag(name, attrs=None, content=None):
    attributes = "".join(f' {key}="{_attr(value)}"' for key, value in (attrs or []))
    if content is None:
        return f"<{name}{attributes}/>"
    return f"<{name}{attributes}>{content}</{name}>"


def _color(value):
    if value is None:
        return None
    text = str(value).lower()
    text = _COLOR_NAMES.get(text, text.lstrip("#"))
    if len(text) == 6:
        return "FF" + text.upper()
    return text.upper()


class Workbook:
    def __init__(self, filename=None, options=None):
        options = options or {}
        self.filename = filename
        self.date_1904 = bool(options.get("date_1904", False))
        self.strings_to_numbers = bool(options.get("strings_to_numbers", False))
        self.strings_to_formulas = bool(options.get("strings_to_formulas", True))
        self.strings_to_urls = bool(options.get("strings_to_urls", True))
        self.nan_inf_to_errors = bool(options.get("nan_inf_to_errors", False))
        self.remove_timezone = bool(options.get("remove_timezone", False))
        self.in_memory = bool(options.get("in_memory", False))
        self.constant_memory = bool(options.get("constant_memory", False))
        if self.constant_memory:
            raise NotImplementedError(
                "constant_memory uses inline strings and is outside the shared-string subset"
            )
        self.max_url_length = max(255, int(options.get("max_url_length", 2079)))
        self.allow_zip64 = bool(options.get("use_zip64", False))
        self.worksheets_objs: list[Worksheet] = []
        self.sheetnames: dict[str, Worksheet] = {}
        self.formats: list[Format] = []
        self._used_formats: list[Format] = []
        self._strings: list[str] = []
        self._closed = False
        self._properties: dict = {}
        self._default = Format({"xf_index": 0}, self)
        self.default_url_format = Format(
            {"font_color": "blue", "underline": 1, "hyperlink": True}, self
        )
        default_date = options.get("default_date_format")
        self.default_date_format = (
            self.add_format({"num_format": default_date}) if default_date else None
        )

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        if exc_type is None:
            self.close()

    def add_worksheet(self, name=None, worksheet_class=None):
        if name is None:
            number = len(self.worksheets_objs) + 1
            name = f"Sheet{number}"
            while name.lower() in self.sheetnames:
                number += 1
                name = f"Sheet{number}"
        if len(name) > 31:
            raise InvalidWorksheetName(
                f"Excel worksheet name {name!r} must be <= 31 chars."
            )
        if re.search(r"[\[\]:*?/\\]", name):
            raise InvalidWorksheetName(
                f"Invalid Excel character '[]:*?/\\' in sheetname {name!r}."
            )
        if name.startswith("'") or name.endswith("'"):
            raise InvalidWorksheetName("Sheet name cannot start or end with an apostrophe.")
        if name.lower() in self.sheetnames:
            raise DuplicateWorksheetName(
                f"Sheetname {name!r}, with case ignored, is already in use."
            )
        cls = worksheet_class or Worksheet
        worksheet = cls(self, name, len(self.worksheets_objs))
        self.worksheets_objs.append(worksheet)
        self.sheetnames[name.lower()] = worksheet
        return worksheet

    def add_format(self, properties=None):
        cell_format = Format(properties, self)
        self.formats.append(cell_format)
        return cell_format

    def get_worksheet_by_name(self, name):
        return self.sheetnames.get(name.lower())

    def worksheets(self):
        return list(self.worksheets_objs)

    def set_properties(self, properties):
        self._properties = dict(properties or {})

    def use_zip64(self):
        self.allow_zip64 = True

    def _add_string(self, value):
        occurrence = len(self._strings)
        self._strings.append(value)
        return occurrence

    def _style_id(self, cell_format):
        if cell_format is None or cell_format is self._default:
            return 0
        if cell_format._style_index is None:
            cell_format._style_index = len(self._used_formats) + 1
            self._used_formats.append(cell_format)
        return cell_format._style_index

    def close(self):
        if self._closed:
            return
        if not self.worksheets_objs:
            self.add_worksheet()
        ids, unique_strings = self._index_strings()
        escaped_strings = escape_xml(unique_strings)
        try:
            destination = self.filename
            if destination is None:
                raise FileCreateError("Workbook filename is required")
            with zipfile.ZipFile(
                destination,
                "w",
                compression=zipfile.ZIP_DEFLATED,
                allowZip64=self.allow_zip64,
            ) as archive:
                self._write_package(archive, ids, unique_strings, escaped_strings)
        except (OSError, zipfile.BadZipFile) as error:
            raise FileCreateError(error) from error
        self._closed = True

    def _index_strings(self):
        return deduplicate(self._strings)

    def _write_package(self, archive, ids, unique_strings, escaped_strings):
        for index, worksheet in enumerate(self.worksheets_objs, 1):
            xml, rels = self._worksheet_xml(worksheet, ids)
            archive.writestr(f"xl/worksheets/sheet{index}.xml", xml)
            if rels:
                archive.writestr(
                    f"xl/worksheets/_rels/sheet{index}.xml.rels", rels
                )
        if self._strings:
            archive.writestr(
                "xl/sharedStrings.xml",
                self._shared_strings_xml(unique_strings, escaped_strings),
            )
        archive.writestr("xl/styles.xml", self._styles_xml())
        archive.writestr("xl/workbook.xml", self._workbook_xml())
        archive.writestr("xl/_rels/workbook.xml.rels", self._workbook_rels())
        archive.writestr("_rels/.rels", self._root_rels())
        archive.writestr("[Content_Types].xml", self._content_types())
        archive.writestr("docProps/core.xml", self._core_xml())
        archive.writestr("docProps/app.xml", self._app_xml())

    def _shared_strings_xml(self, strings, escaped):
        pieces = [
            XML,
            f'<sst xmlns="{MAIN_NS}" count="{len(self._strings)}" uniqueCount="{len(strings)}">',
        ]
        for original, value in zip(strings, escaped, strict=True):
            preserve = bool(original) and (original[0].isspace() or original[-1].isspace())
            attr = ' xml:space="preserve"' if preserve else ""
            pieces.append(f"<si><t{attr}>{value}</t></si>")
        pieces.append("</sst>")
        return "".join(pieces)

    def _worksheet_xml(self, worksheet, string_ids):
        rel_entries = []
        hyperlink_entries = []
        for hyperlink in worksheet.hyperlinks:
            target, location = worksheet._external_url(hyperlink["url"])
            cell = xl_rowcol_to_cell(hyperlink["row"], hyperlink["col"])
            attrs = [("ref", cell)]
            if location is not None:
                attrs.append(("location", location))
            else:
                rel_id = f"rId{len(rel_entries) + 1}"
                attrs.append(("r:id", rel_id))
                rel_entries.append((rel_id, target))
            if hyperlink["tip"]:
                attrs.append(("tooltip", hyperlink["tip"]))
            hyperlink_entries.append(_tag("hyperlink", attrs))

        if worksheet.dim_rowmin is None:
            dimension = "A1"
        else:
            dimension = xl_range(
                worksheet.dim_rowmin,
                worksheet.dim_colmin,
                worksheet.dim_rowmax,
                worksheet.dim_colmax,
            )
        pieces = [
            XML,
            f'<worksheet xmlns="{MAIN_NS}" xmlns:r="{REL_NS}">',
            _tag("dimension", [("ref", dimension)]),
            self._sheet_views_xml(worksheet),
            '<sheetFormatPr defaultRowHeight="15"/>',
        ]
        columns = self._columns_xml(worksheet)
        if columns:
            pieces.append(columns)
        pieces.append("<sheetData>")
        rows = sorted(set(worksheet.cells) | set(worksheet.row_options))
        column_names = {}
        for row in rows:
            row_cells = sorted(worksheet.cells.get(row, {}).items())
            attrs = [("r", row + 1)]
            if row_cells:
                attrs.append(("spans", f"{row_cells[0][0] + 1}:{row_cells[-1][0] + 1}"))
            row_format = None
            if row in worksheet.row_options:
                height, row_format, hidden, level, collapsed = worksheet.row_options[row]
                if height != 15:
                    attrs += [("ht", f"{height:g}"), ("customHeight", 1)]
                if row_format:
                    attrs += [("s", self._style_id(row_format)), ("customFormat", 1)]
                if hidden:
                    attrs.append(("hidden", 1))
                if level:
                    attrs.append(("outlineLevel", level))
                if collapsed:
                    attrs.append(("collapsed", 1))
            cell_pieces = []
            for col, cell in row_cells:
                column_name = column_names.get(col)
                if column_name is None:
                    column_name = xl_rowcol_to_cell(0, col)[:-1]
                    column_names[col] = column_name
                cell_pieces.append(
                    self._cell_xml(
                        worksheet,
                        row,
                        col,
                        cell,
                        row_format,
                        string_ids,
                        column_name + str(row + 1),
                    )
                )
            content = "".join(cell_pieces)
            attributes = "".join(f' {key}="{value}"' for key, value in attrs)
            pieces.append(f"<row{attributes}>{content}</row>")
        pieces.append("</sheetData>")
        if worksheet.autofilter_ref:
            pieces.append(_tag("autoFilter", [("ref", worksheet.autofilter_ref)]))
        if worksheet.merges:
            merge_xml = "".join(
                _tag("mergeCell", [("ref", xl_range(*cell_range))])
                for cell_range in worksheet.merges
            )
            pieces.append(
                _tag("mergeCells", [("count", len(worksheet.merges))], merge_xml)
            )
        if hyperlink_entries:
            pieces.append(_tag("hyperlinks", content="".join(hyperlink_entries)))
        pieces.append(
            '<pageMargins left="0.7" right="0.7" top="0.75" '
            'bottom="0.75" header="0.3" footer="0.3"/>'
        )
        pieces.append("</worksheet>")
        rels = ""
        if rel_entries:
            relationships = "".join(
                _tag(
                    "Relationship",
                    [
                        ("Id", rel_id),
                        ("Type", f"{REL_NS}/hyperlink"),
                        ("Target", target),
                        ("TargetMode", "External"),
                    ],
                )
                for rel_id, target in rel_entries
            )
            rels = XML + f'<Relationships xmlns="{PKG_REL_NS}">{relationships}</Relationships>'
        return "".join(pieces), rels

    def _sheet_views_xml(self, worksheet):
        attrs = [("workbookViewId", 0)]
        if worksheet.selected:
            attrs.insert(0, ("tabSelected", 1))
        content = ""
        if worksheet.panes:
            row, col, top_row, left_col = worksheet.panes
            pane_attrs = []
            if col:
                pane_attrs.append(("xSplit", col))
            if row:
                pane_attrs.append(("ySplit", row))
            pane_attrs.append(("topLeftCell", xl_rowcol_to_cell(top_row, left_col)))
            pane_attrs.append(
                ("activePane", "bottomRight" if row and col else "bottomLeft" if row else "topRight")
            )
            pane_attrs.append(("state", "frozen"))
            content = _tag("pane", pane_attrs)
        return _tag("sheetViews", content=_tag("sheetView", attrs, content))

    def _columns_xml(self, worksheet):
        if not worksheet.col_options:
            return ""
        pieces = []
        columns = sorted(worksheet.col_options)
        start = previous = columns[0]
        record = worksheet.col_options[start]
        groups = []
        for column in columns[1:]:
            current = worksheet.col_options[column]
            if column == previous + 1 and current == record:
                previous = column
                continue
            groups.append((start, previous, record))
            start = previous = column
            record = current
        groups.append((start, previous, record))
        for first, last, (width, cell_format, hidden, level, collapsed) in groups:
            custom = 1
            if width is None:
                width = 0 if hidden else 8.43
                custom = int(hidden)
            if width > 0:
                if width < 1:
                    width = int(int(width * 12 + 0.5) / 7 * 256) / 256
                else:
                    width = int((int(width * 7 + 0.5) + 5) / 7 * 256) / 256
            attrs = [("min", first + 1), ("max", last + 1), ("width", f"{width:.16g}")]
            if cell_format:
                attrs.append(("style", self._style_id(cell_format)))
            if hidden:
                attrs.append(("hidden", 1))
            if custom:
                attrs.append(("customWidth", 1))
            if level:
                attrs.append(("outlineLevel", level))
            if collapsed:
                attrs.append(("collapsed", 1))
            pieces.append(_tag("col", attrs))
        return _tag("cols", content="".join(pieces))

    def _cell_xml(self, worksheet, row, col, cell, row_format, ids, reference=None):
        cell_format = cell.cell_format or row_format
        if cell_format is None and col in worksheet.col_options:
            cell_format = worksheet.col_options[col][1]
        reference = reference or xl_rowcol_to_cell(row, col)
        attrs = f' r="{reference}"'
        style = self._style_id(cell_format)
        if style:
            attrs += f' s="{style}"'
        if cell.kind == "blank":
            return f"<c{attrs}/>"
        if cell.kind == "string":
            value = int(ids[int(cell.value)])
            return f'<c{attrs} t="s"><v>{value}</v></c>'
        elif cell.kind == "boolean":
            value = "1" if cell.value else "0"
            return f'<c{attrs} t="b"><v>{value}</v></c>'
        elif cell.kind == "formula":
            formula, value = cell.value
            content = f"<f>{html.escape(str(formula), quote=False)}</f>"
            if isinstance(value, bool):
                attrs += ' t="b"'
                cached = "1" if value else "0"
            elif isinstance(value, str):
                attrs += ' t="e"' if value.startswith("#") else ' t="str"'
                cached = html.escape(value, quote=False)
            else:
                cached = _number(value)
            content += f"<v>{cached}</v>"
        else:
            content = f"<v>{_number(cell.value)}</v>"
        return f"<c{attrs}>{content}</c>"

    def _styles_xml(self):
        formats = [self._default] + self._used_formats
        font_records = []
        fill_records = [None, "gray125"]
        border_records = [None]
        num_formats = []
        format_refs = []

        def identify(records, record):
            if record not in records:
                records.append(record)
            return records.index(record)

        for cell_format in formats:
            props = cell_format.properties
            font = (
                bool(props.get("bold")),
                bool(props.get("italic")),
                props.get("underline"),
                props.get("font_name", "Calibri"),
                props.get("font_size", 11),
                _color(props.get("font_color")),
                bool(props.get("font_strikeout")),
            )
            font_id = identify(font_records, font)
            fill = (
                _color(props.get("bg_color") or props.get("fg_color")),
                props.get("pattern", 1 if props.get("bg_color") or props.get("fg_color") else 0),
            )
            fill_id = 0 if not fill[0] and not fill[1] else identify(fill_records, fill)
            border = (
                props.get("left", props.get("border", 0)),
                props.get("right", props.get("border", 0)),
                props.get("top", props.get("border", 0)),
                props.get("bottom", props.get("border", 0)),
                _color(props.get("border_color")),
            )
            border_id = 0 if not any(border[:4]) else identify(border_records, border)
            num_format = props.get("num_format")
            if num_format:
                if num_format not in num_formats:
                    num_formats.append(num_format)
                num_fmt_id = 164 + num_formats.index(num_format)
            else:
                num_fmt_id = 0
            format_refs.append((font_id, fill_id, border_id, num_fmt_id, props))

        pieces = [XML, f'<styleSheet xmlns="{MAIN_NS}">']
        if num_formats:
            pieces.append(
                _tag(
                    "numFmts",
                    [("count", len(num_formats))],
                    "".join(
                        _tag("numFmt", [("numFmtId", 164 + index), ("formatCode", value)])
                        for index, value in enumerate(num_formats)
                    ),
                )
            )
        pieces.append(
            _tag(
                "fonts",
                [("count", len(font_records))],
                "".join(self._font_xml(font) for font in font_records),
            )
        )
        pieces.append(
            _tag(
                "fills",
                [("count", len(fill_records))],
                "".join(self._fill_xml(fill) for fill in fill_records),
            )
        )
        pieces.append(
            _tag(
                "borders",
                [("count", len(border_records))],
                "".join(self._border_xml(border) for border in border_records),
            )
        )
        pieces.append(
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" '
            'fillId="0" borderId="0"/></cellStyleXfs>'
        )
        xfs = []
        for font_id, fill_id, border_id, num_fmt_id, props in format_refs:
            attrs = [
                ("numFmtId", num_fmt_id),
                ("fontId", font_id),
                ("fillId", fill_id),
                ("borderId", border_id),
                ("xfId", 0),
            ]
            if num_fmt_id:
                attrs.append(("applyNumberFormat", 1))
            if font_id:
                attrs.append(("applyFont", 1))
            if fill_id:
                attrs.append(("applyFill", 1))
            if border_id:
                attrs.append(("applyBorder", 1))
            alignment = self._alignment_xml(props)
            protection = self._protection_xml(props)
            if alignment:
                attrs.append(("applyAlignment", 1))
            if protection:
                attrs.append(("applyProtection", 1))
            xfs.append(_tag("xf", attrs, alignment + protection if alignment or protection else None))
        pieces.append(_tag("cellXfs", [("count", len(xfs))], "".join(xfs)))
        pieces.append(
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" '
            'builtinId="0"/></cellStyles><dxfs count="0"/>'
            '<tableStyles count="0" defaultTableStyle="TableStyleMedium9" '
            'defaultPivotStyle="PivotStyleLight16"/>'
        )
        pieces.append("</styleSheet>")
        return "".join(pieces)

    def _font_xml(self, font):
        bold, italic, underline, name, size, color, strike = font
        content = ""
        if bold:
            content += "<b/>"
        if italic:
            content += "<i/>"
        if strike:
            content += "<strike/>"
        if underline:
            value = {2: "double", 33: "singleAccounting", 34: "doubleAccounting"}.get(underline)
            content += _tag("u", [("val", value)]) if value else "<u/>"
        content += _tag("sz", [("val", size)])
        content += _tag("color", [("rgb", color)]) if color else '<color theme="1"/>'
        content += _tag("name", [("val", name)])
        content += '<family val="2"/>'
        if name == "Calibri":
            content += '<scheme val="minor"/>'
        return _tag("font", content=content)

    def _fill_xml(self, fill):
        if fill is None:
            return "<fill><patternFill patternType=\"none\"/></fill>"
        if fill == "gray125":
            return "<fill><patternFill patternType=\"gray125\"/></fill>"
        color, pattern = fill
        pattern_name = "solid" if pattern == 1 else {
            2: "mediumGray", 3: "darkGray", 4: "lightGray",
        }.get(pattern, "solid")
        return (
            f'<fill><patternFill patternType="{pattern_name}">'
            f'<fgColor rgb="{color}"/><bgColor indexed="64"/>'
            "</patternFill></fill>"
        )

    def _border_xml(self, border):
        if border is None:
            return "<border><left/><right/><top/><bottom/><diagonal/></border>"
        left, right, top, bottom, color = border

        def side(name, style):
            if not style:
                return _tag(name)
            child = _tag("color", [("rgb", color)]) if color else '<color auto="1"/>'
            return _tag(name, [("style", _BORDER_STYLES.get(style, "thin"))], child)

        return _tag(
            "border",
            content=side("left", left)
            + side("right", right)
            + side("top", top)
            + side("bottom", bottom)
            + "<diagonal/>",
        )

    def _alignment_xml(self, props):
        attrs = []
        horizontal = {
            "left": "left", "center": "center", "right": "right",
            "fill": "fill", "justify": "justify", "center_across": "centerContinuous",
            "distributed": "distributed",
        }.get(props.get("align"))
        vertical = {
            "top": "top", "vcenter": "center", "bottom": "bottom",
            "vjustify": "justify", "vdistributed": "distributed",
        }.get(props.get("valign"))
        if horizontal:
            attrs.append(("horizontal", horizontal))
        if vertical:
            attrs.append(("vertical", vertical))
        if props.get("text_wrap"):
            attrs.append(("wrapText", 1))
        if props.get("rotation"):
            attrs.append(("textRotation", props["rotation"]))
        if props.get("indent"):
            attrs.append(("indent", props["indent"]))
        if props.get("shrink"):
            attrs.append(("shrinkToFit", 1))
        return _tag("alignment", attrs) if attrs else ""

    def _protection_xml(self, props):
        attrs = []
        if props.get("locked") is False:
            attrs.append(("locked", 0))
        if props.get("hidden"):
            attrs.append(("hidden", 1))
        return _tag("protection", attrs) if attrs else ""

    def _workbook_xml(self):
        sheets = ""
        for index, worksheet in enumerate(self.worksheets_objs, 1):
            attrs = [("name", worksheet.name), ("sheetId", index)]
            if worksheet.hidden:
                attrs.append(("state", "hidden"))
            attrs.append(("r:id", f"rId{index}"))
            sheets += _tag("sheet", attrs)
        defined = []
        for index, worksheet in enumerate(self.worksheets_objs):
            if worksheet.autofilter_ref:
                sheet_name = worksheet.name.replace("'", "''")
                defined.append(
                    _tag(
                        "definedName",
                        [("name", "_xlnm._FilterDatabase"), ("localSheetId", index), ("hidden", 1)],
                        f"'{html.escape(sheet_name)}'!{_absolute_ref(worksheet.autofilter_ref)}",
                    )
                )
        workbook_pr = [("defaultThemeVersion", 124226)]
        if self.date_1904:
            workbook_pr.append(("date1904", 1))
        pieces = [
            XML,
            f'<workbook xmlns="{MAIN_NS}" xmlns:r="{REL_NS}">',
            '<fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>',
            _tag("workbookPr", workbook_pr),
            '<bookViews><workbookView xWindow="240" yWindow="15" '
            'windowWidth="16095" windowHeight="9660"/></bookViews>',
            _tag("sheets", content=sheets),
        ]
        if defined:
            pieces.append(_tag("definedNames", content="".join(defined)))
        pieces += ['<calcPr calcId="124519" fullCalcOnLoad="1"/>', "</workbook>"]
        return "".join(pieces)

    def _workbook_rels(self):
        entries = []
        for index in range(1, len(self.worksheets_objs) + 1):
            entries.append(
                _tag(
                    "Relationship",
                    [
                        ("Id", f"rId{index}"),
                        ("Type", f"{REL_NS}/worksheet"),
                        ("Target", f"worksheets/sheet{index}.xml"),
                    ],
                )
            )
        next_id = len(entries) + 1
        entries.append(
            _tag(
                "Relationship",
                [("Id", f"rId{next_id}"), ("Type", f"{REL_NS}/styles"), ("Target", "styles.xml")],
            )
        )
        if self._strings:
            entries.append(
                _tag(
                    "Relationship",
                    [
                        ("Id", f"rId{next_id + 1}"),
                        ("Type", f"{REL_NS}/sharedStrings"),
                        ("Target", "sharedStrings.xml"),
                    ],
                )
            )
        return XML + f'<Relationships xmlns="{PKG_REL_NS}">{"".join(entries)}</Relationships>'

    def _root_rels(self):
        relationships = "".join(
            [
                _tag("Relationship", [("Id", "rId1"), ("Type", f"{REL_NS}/officeDocument"), ("Target", "xl/workbook.xml")]),
                _tag("Relationship", [("Id", "rId2"), ("Type", f"{PKG_REL_NS}/metadata/core-properties"), ("Target", "docProps/core.xml")]),
                _tag("Relationship", [("Id", "rId3"), ("Type", f"{REL_NS}/extended-properties"), ("Target", "docProps/app.xml")]),
            ]
        )
        return XML + f'<Relationships xmlns="{PKG_REL_NS}">{relationships}</Relationships>'

    def _content_types(self):
        defaults = (
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
        )
        overrides = [
            ("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml"),
            ("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
            ("/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"),
            ("/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"),
        ]
        if self._strings:
            overrides.append(
                (
                    "/xl/sharedStrings.xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
                )
            )
        overrides.extend(
            (f"/xl/worksheets/sheet{index}.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
            for index in range(1, len(self.worksheets_objs) + 1)
        )
        content = defaults + "".join(
            _tag("Override", [("PartName", name), ("ContentType", content_type)])
            for name, content_type in overrides
        )
        return XML + f'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">{content}</Types>'

    def _core_xml(self):
        properties = self._properties
        created = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        creator = html.escape(str(properties.get("author", "")))
        title = properties.get("title")
        subject = properties.get("subject")
        extra = _tag("dc:title", content=html.escape(str(title))) if title else ""
        extra += _tag("dc:subject", content=html.escape(str(subject))) if subject else ""
        return (
            XML
            + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" '
            'xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f"<dc:creator>{creator}</dc:creator><cp:lastModifiedBy>{creator}</cp:lastModifiedBy>"
            + extra
            + f'<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{created}</dcterms:modified>'
            "</cp:coreProperties>"
        )

    def _app_xml(self):
        names = "".join(_tag("vt:lpstr", content=html.escape(ws.name)) for ws in self.worksheets_objs)
        return (
            XML
            + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            "<Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>"
            '<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr>'
            f'</vt:variant><vt:variant><vt:i4>{len(self.worksheets_objs)}</vt:i4></vt:variant></vt:vector></HeadingPairs>'
            f'<TitlesOfParts><vt:vector size="{len(self.worksheets_objs)}" baseType="lpstr">{names}</vt:vector></TitlesOfParts>'
            "<Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>"
            "<HyperlinksChanged>false</HyperlinksChanged><AppVersion>12.0000</AppVersion></Properties>"
        )


def _number(value):
    if isinstance(value, bool):
        return "1" if value else "0"
    if isinstance(value, int):
        return str(value)
    return f"{float(value):.16g}"


def _absolute_ref(reference):
    parts = reference.split(":")
    absolute = []
    for part in parts:
        match = re.match(r"([A-Z]+)(\d+)", part)
        absolute.append(f"${match.group(1)}${match.group(2)}")
    return ":".join(absolute)
