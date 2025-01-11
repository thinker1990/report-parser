from pandas import DataFrame
from xlsxwriter.utility import xl_cell_to_rowcol

RANGE_SPLITTER = ":"


def range_to_slice(range_notation: str) -> tuple[slice, slice]:
    start, stop = range_notation.split(RANGE_SPLITTER)
    start_row, start_col = xl_cell_to_rowcol(start)
    stop_row, stop_col = xl_cell_to_rowcol(stop)
    return slice(start_row, stop_row + 1), slice(start_col, stop_col + 1)


def parse_row_range(range_notation: str) -> slice:
    start_str, end_str = range_notation.split(RANGE_SPLITTER)
    start = int(start_str) if start_str else 0
    end = int(end_str) if end_str else None
    return slice(start - 1, end)


def parse_cells(mapping: dict, sheet: DataFrame) -> dict:
    result = {}
    for name, location in mapping.items():
        row_index, col_index = xl_cell_to_rowcol(location)
        result[name] = sheet.iloc[row_index, col_index]
    return result


def parse_ranges(mapping: dict, sheet: DataFrame) -> dict:
    result = {}
    for name, range_notation in mapping.items():
        row_slice, col_slice = range_to_slice(range_notation)
        range_values = sheet.iloc[row_slice, col_slice].values
        result[name] = range_values.flatten().tolist()
    return result


def parse_row_components(row_index: int, mapping: dict, sheet: DataFrame) -> dict:
    result = {}
    for name, column in mapping.items():
        if RANGE_SPLITTER not in column:
            _, col_index = xl_cell_to_rowcol(column)
            result[name] = sheet.iloc[row_index, col_index]
        else:
            _, col_slice = range_to_slice(column)
            range_values = sheet.iloc[[row_index], col_slice].values
            result[name] = range_values.flatten().tolist()
    return result


def parse_repeat_rows(mapping: dict, sheet: DataFrame) -> list:
    result = []
    for row_index in sheet.index[parse_row_range(mapping["range"])]:
        row = parse_row_components(row_index, mapping["components"], sheet)
        result.append(row)
    return result
