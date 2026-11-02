import pytest
from openpyxl import Workbook

from openpyxl_chartkit import wide_to_long


@pytest.fixture
def workbook_with_data():
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "data"

    worksheet.append([
        "SampleNum",
        "Strain %",
        "G' MPa",
        'G" MPa',
        "tan-delta",
    ])
    worksheet.append([1, 0.1, 10.0, 5.0, 0.5])
    worksheet.append([1, 0.2, 20.0, 10.0, 0.5])
    worksheet.append([2, 0.1, 30.0, 15.0, 0.5])
    worksheet.append([2, 0.2, 40.0, None, 0.6])

    return workbook, worksheet


def test_wide_to_long_converts_data(workbook_with_data):
    workbook, source = workbook_with_data

    result = wide_to_long(
        source=source,
        id_columns=("SampleNum", "Strain %"),
        value_columns=("G' MPa", 'G" MPa', "tan-delta"),
    )

    assert result.title == "chart_data"
    assert list(result.values)[0] == (
        "SampleNum", "Strain %", "Metric", "Value"
    )

    assert list(result.values)[1] == (
        1, 0.1, "G' MPa", 10.0
    )
    assert list(result.values)[2] == (
        1, 0.1, 'G" MPa', 5.0
    )
    assert list(result.values)[3] == (
        1, 0.1, "tan-delta", 0.5
    )

    assert result.max_row == 12


def test_wide_to_long_skips_none_values(workbook_with_data):
    _, source = workbook_with_data

    result = wide_to_long(
        source=source,
        id_columns=("SampleNum", "Strain %"),
        value_columns=("G' MPa", 'G" MPa', "tan-delta"),
    )

    rows = list(result.iter_rows(min_row=2, values_only=True))

    assert len(rows) == 11
    assert not any(
        row[2] == 'G" MPa' and row[0] == 2 and row[1] == 0.2
        for row in rows
    )


def test_wide_to_long_preserves_source_worksheet(workbook_with_data):
    _, source = workbook_with_data
    original_values = list(source.values)

    wide_to_long(
        source=source,
        id_columns=("SampleNum", "Strain %"),
        value_columns=("G' MPa", 'G" MPa', "tan-delta"),
    )

    assert list(source.values) == original_values


def test_wide_to_long_supports_multiple_id_columns(workbook_with_data):
    _, source = workbook_with_data

    result = wide_to_long(
        source=source,
        id_columns=("SampleNum", "Strain %"),
        value_columns=("G' MPa",),
    )

    assert list(result.values)[1] == (
        1, 0.1, "G' MPa", 10.0
    )


def test_wide_to_long_raises_error_for_missing_column(workbook_with_data):
    _, source = workbook_with_data

    with pytest.raises(ValueError, match="Missing columns"):
        wide_to_long(
            source=source,
            id_columns=("SampleNum",),
            value_columns=("UnknownColumn",),
        )


def test_wide_to_long_raises_error_for_duplicate_headers():
    workbook = Workbook()
    source = workbook.active
    source.append(["SampleNum", "Value", "Value"])
    source.append([1, 10, 20])

    with pytest.raises(ValueError, match="duplicate headers"):
        wide_to_long(
            source=source,
            id_columns=("SampleNum",),
            value_columns=("Value",),
        )


def test_wide_to_long_replaces_existing_output_sheet(workbook_with_data):
    workbook, source = workbook_with_data

    existing = workbook.create_sheet("chart_data")
    existing["A1"] = "old data"

    result = wide_to_long(
        source=source,
        id_columns=("SampleNum",),
        value_columns=("G' MPa",),
    )

    assert result is workbook["chart_data"]
    assert result["A1"].value == "SampleNum"
    assert existing not in workbook.worksheets


def test_wide_to_long_rejects_source_as_output_sheet(workbook_with_data):
    _, source = workbook_with_data

    with pytest.raises(ValueError, match="cannot be the source"):
        wide_to_long(
            source=source,
            id_columns=("SampleNum",),
            value_columns=("G' MPa",),
            output_sheet_name="data",
        )


def test_wide_to_long_sorts_by_specified_columns(workbook_with_data):
    _, source = workbook_with_data

    result = wide_to_long(
        source=source,
        id_columns=("SampleNum", "Strain %"),
        value_columns=("G' MPa", 'G" MPa', "tan-delta"),
        sort_by=("SampleNum", "Metric", "Strain %"),
    )

    rows = list(result.iter_rows(min_row=2, values_only=True))

    assert rows[:2] == [
        (1, 0.1, "G' MPa", 10.0),
        (1, 0.2, "G' MPa", 20.0),
    ]