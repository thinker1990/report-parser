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


def extract_cell_value(location: str, sheet: DataFrame, row_index: int = None):
    """Extract a single cell value from the sheet.
    
    Args:
        location: Cell location (e.g., "A1")
        sheet: DataFrame to extract from
        row_index: If provided, uses this row index instead of the one from location
    
    Returns:
        The cell value
    """
    loc_row_index, col_index = xl_cell_to_rowcol(location)
    actual_row = row_index if row_index is not None else loc_row_index
    return sheet.iloc[actual_row, col_index]


def extract_range_values(range_notation: str, sheet: DataFrame, row_index: int = None) -> list:
    """Extract values from a range in the sheet.
    
    Args:
        range_notation: Cell range notation (e.g., "A1:B2")
        sheet: DataFrame to extract from
        row_index: If provided, extracts only from this specific row
    
    Returns:
        Flattened list of values from the range
    """
    row_slice, col_slice = range_to_slice(range_notation)
    if row_index is not None:
        range_values = sheet.iloc[[row_index], col_slice].values
    else:
        range_values = sheet.iloc[row_slice, col_slice].values
    return range_values.flatten().tolist()


def parse_cells(mapping: dict, sheet: DataFrame) -> dict:
    result = {}
    for name, location in mapping.items():
        result[name] = extract_cell_value(location, sheet)
    return result


def parse_ranges(mapping: dict, sheet: DataFrame) -> dict:
    result = {}
    for name, range_notation in mapping.items():
        result[name] = extract_range_values(range_notation, sheet)
    return result


def parse_row_components(row_index: int, mapping: dict, sheet: DataFrame) -> dict:
    result = {}
    for name, column in mapping.items():
        if RANGE_SPLITTER not in column:
            result[name] = extract_cell_value(column, sheet, row_index)
        else:
            result[name] = extract_range_values(column, sheet, row_index)
    return result


def parse_repeat_rows(mapping: dict, sheet: DataFrame) -> list:
    result = []
    for row_index in sheet.index[parse_row_range(mapping["range"])]:
        row = parse_row_components(row_index, mapping["components"], sheet)
        result.append(row)
    return result
