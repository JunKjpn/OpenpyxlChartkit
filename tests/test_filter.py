import pytest
from openpyxl import Workbook

from openpyxl_chartkit import filter_rows


@pytest.fixture
def workbook_with_data():
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "data"

    worksheet.append(["SampleNum", "Strain %", "Stress MPa"])
    worksheet.append([1, 0.1, 10.0])
    worksheet.append([1, 0.2, 20.0])
    worksheet.append([2, 0.1, 30.0])
    worksheet.append([2, 0.2, 40.0])

    return workbook, worksheet


def test_filter_rows_extracts_matching_rows(workbook_with_data):
    _, source = workbook_with_data

    result = filter_rows(
        source=source,
        column="SampleNum",
        value=2,
    )

    assert result.title == "filtered_data"
    assert list(result.values) == [
        ("SampleNum", "Strain %", "Stress MPa"),
        (2, 0.1, 30.0),
        (2, 0.2, 40.0),
    ]


def test_filter_rows_preserves_row_order(workbook_with_data):
    _, source = workbook_with_data

    result = filter_rows(
        source=source,
        column="SampleNum",
        value=1,
    )

    rows = list(result.iter_rows(min_row=2, values_only=True))

    assert rows == [
        (1, 0.1, 10.0),
        (1, 0.2, 20.0),
    ]


def test_filter_rows_returns_header_only_when_no_match(workbook_with_data):
    _, source = workbook_with_data

    result = filter_rows(
        source=source,
        column="SampleNum",
        value=99,
    )

    assert list(result.values) == [
        ("SampleNum", "Strain %", "Stress MPa"),
    ]


def test_filter_rows_preserves_source_worksheet(workbook_with_data):
    _, source = workbook_with_data
    original_values = list(source.values)

    filter_rows(
        source=source,
        column="SampleNum",
        value=2,
    )

    assert list(source.values) == original_values


def test_filter_rows_raises_error_for_missing_column(workbook_with_data):
    _, source = workbook_with_data

    with pytest.raises(ValueError, match="not found"):
        filter_rows(
            source=source,
            column="UnknownColumn",
            value=2,
        )


def test_filter_rows_raises_error_for_duplicate_headers():
    workbook = Workbook()
    source = workbook.active
    source.append(["SampleNum", "Value", "Value"])
    source.append([1, 10, 20])

    with pytest.raises(ValueError, match="duplicate headers"):
        filter_rows(
            source=source,
            column="SampleNum",
            value=1,
        )


def test_filter_rows_replaces_existing_output_sheet(workbook_with_data):
    workbook, source = workbook_with_data

    existing = workbook.create_sheet("filtered_data")
    existing["A1"] = "old data"

    result = filter_rows(
        source=source,
        column="SampleNum",
        value=2,
    )

    assert result is workbook["filtered_data"]
    assert result["A1"].value == "SampleNum"
    assert existing not in workbook.worksheets


def test_filter_rows_rejects_source_as_output_sheet(workbook_with_data):
    _, source = workbook_with_data

    with pytest.raises(ValueError, match="cannot be the source"):
        filter_rows(
            source=source,
            column="SampleNum",
            value=2,
            output_sheet_name="data",
        )
