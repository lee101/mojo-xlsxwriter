"""mojo-xlsxwriter against XlsxWriter on identical workbook data."""

from __future__ import annotations

import gc
import math
import os
import platform
import random
import sys
import time
import zipfile
from io import BytesIO
from xml.etree import ElementTree as ET

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "python"))

import xlsxwriter  # noqa: E402
from xlsxwriter.xmlwriter import XMLwriter  # noqa: E402

import mojo_xlsxwriter as mxw  # noqa: E402
from mojo_xlsxwriter._lib import deduplicate, escape_xml  # noqa: E402


def best_time(function, repeats=3):
    function()
    best = math.inf
    result = None
    for _ in range(repeats):
        gc.collect()
        start = time.perf_counter()
        result = function()
        best = min(best, time.perf_counter() - start)
    return best, result


def cpu_name():
    try:
        with open("/proc/cpuinfo", encoding="utf-8") as source:
            for line in source:
                if line.startswith("model name"):
                    return line.split(":", 1)[1].strip()
    except OSError:
        pass
    return platform.processor() or "unknown CPU"


def reference_strings(values):
    table = {}
    ids = []
    unique = []
    for value in values:
        identifier = table.get(value)
        if identifier is None:
            identifier = len(unique)
            table[value] = identifier
            unique.append(value)
        ids.append(identifier)
    writer = XMLwriter()
    escaped = [
        writer._escape_data(writer._escape_control_characters(value))
        for value in unique
    ]
    return ids, unique, escaped


def mojo_strings(values):
    ids, unique = deduplicate(values)
    return ids.tolist(), unique, escape_xml(unique)


def reference_escape(values):
    writer = XMLwriter()
    return [
        writer._escape_data(writer._escape_control_characters(value))
        for value in values
    ]


def string_workbook(workbook_class, rows, vocabulary):
    destination = BytesIO()
    workbook = workbook_class(destination, {"in_memory": True})
    sheet = workbook.add_worksheet("Strings")
    for row in range(rows):
        sheet.write_row(
            row,
            0,
            (
                vocabulary[row % len(vocabulary)],
                vocabulary[(row * 17) % len(vocabulary)],
                vocabulary[(row * 97) % len(vocabulary)],
            ),
        )
    workbook.close()
    return destination.getvalue()


def mixed_workbook(workbook_class, rows, vocabulary):
    destination = BytesIO()
    workbook = workbook_class(destination, {"in_memory": True})
    sheet = workbook.add_worksheet("Data")
    money = workbook.add_format({"num_format": "$#,##0.00"})
    sheet.write_row(0, 0, ["key", "value", "flag", "formula"], money)
    for row in range(1, rows + 1):
        sheet.write(row, 0, vocabulary[row % len(vocabulary)])
        sheet.write_number(row, 1, row * 0.125, money)
        sheet.write_boolean(row, 2, row & 1)
        sheet.write_formula(row, 3, f"=B{row + 1}*2", money, row * 0.25)
    sheet.autofilter(0, 0, rows, 3)
    sheet.freeze_panes(1, 0)
    workbook.close()
    return destination.getvalue()


def workbook_signature(data):
    with zipfile.ZipFile(BytesIO(data)) as archive:
        sheet = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        shared = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    namespace = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    return (
        sheet.find("m:dimension", namespace).attrib["ref"],
        len(sheet.findall(".//m:c", namespace)),
        int(shared.attrib["count"]),
        int(shared.attrib["uniqueCount"]),
    )


def print_row(name, mojo_time, reference_time):
    ratio = reference_time / mojo_time
    result = "faster" if ratio >= 1 else "slower"
    print(
        f"| {name} | {mojo_time * 1e3:.2f} ms | "
        f"{reference_time * 1e3:.2f} ms | {ratio:.2f}x | {result} |"
    )


def main():
    rng = random.Random(23)
    vocabulary = [
        f"item-{index:05d}-"
        + "".join(rng.choices("abcdefβ中&<>", k=20))
        + ("_x0000_\x01" if index % 31 == 0 else "")
        for index in range(20_000)
    ]
    values = [
        vocabulary[(index * 104_729) % len(vocabulary)]
        for index in range(600_000)
    ]
    unique_values = [
        f"unique-{index:06d}-&<\x01_x0000_-β中"
        for index in range(300_000)
    ]

    print(f"Machine: {cpu_name()}; {os.cpu_count()} logical CPUs")
    print(
        f"Software: Python {platform.python_version()}, "
        f"XlsxWriter {xlsxwriter.__version__}; best of 3 warm runs"
    )
    print()
    print("| case | mojo-xlsxwriter | XlsxWriter | ratio | result |")
    print("|---|---:|---:|---:|---|")

    mojo_time, mojo_result = best_time(lambda: escape_xml(unique_values))
    reference_time, reference_result = best_time(lambda: reference_escape(unique_values))
    if mojo_result != reference_result:
        raise AssertionError("XML escape benchmark outputs differ")
    print_row("XML escape (300k unique)", mojo_time, reference_time)

    mojo_time, mojo_result = best_time(lambda: mojo_strings(values))
    reference_time, reference_result = best_time(lambda: reference_strings(values))
    if mojo_result != reference_result:
        raise AssertionError("shared-string kernel benchmark outputs differ")
    print_row("dedup + escape (600k strings)", mojo_time, reference_time)

    mojo_time, mojo_result = best_time(
        lambda: string_workbook(mxw.Workbook, 150_000, vocabulary)
    )
    reference_time, reference_result = best_time(
        lambda: string_workbook(xlsxwriter.Workbook, 150_000, vocabulary)
    )
    if workbook_signature(mojo_result) != workbook_signature(reference_result):
        raise AssertionError("string workbook benchmark outputs differ")
    print_row("XLSX: 150k x 3 strings", mojo_time, reference_time)

    mojo_time, mojo_result = best_time(
        lambda: mixed_workbook(mxw.Workbook, 100_000, vocabulary)
    )
    reference_time, reference_result = best_time(
        lambda: mixed_workbook(xlsxwriter.Workbook, 100_000, vocabulary)
    )
    if workbook_signature(mojo_result) != workbook_signature(reference_result):
        raise AssertionError("mixed workbook benchmark outputs differ")
    print_row("XLSX: 100k mixed rows", mojo_time, reference_time)


if __name__ == "__main__":
    main()
