from pandas import DataFrame
from xlsxwriter.utility import xl_cell_to_rowcol


RANGE_SPLITTER = ":"

def parse_range(range: str) -> tuple:
    """
    Parses a string representing a range of Excel cells and returns a tuple of slices
    representing the row and column indices.
    
    Args:
        range (str): A string representing the range of Excel cells. The range can be a 
                     single cell (e.g., 'A1') or a range of cells (e.g., 'A1:B2').
    
    Returns:
        tuple: A tuple containing two slices. The first slice represents the row indices,
               and the second slice represents the column indices.
    
    Raises:
        ValueError: If the range string is not properly formatted or contains invalid cell references.

    """
    if RANGE_SPLITTER not in range:
        return xl_cell_to_rowcol(range)
    
    start, stop = range.split(RANGE_SPLITTER)
    (start_row, start_col) = xl_cell_to_rowcol(start)
    (stop_row, stop_col) = xl_cell_to_rowcol(stop)
    return slice(start_row, stop_row + 1), slice(start_col, stop_col + 1)


def parse_row_range(range: str) -> slice:
    """
    Parses a string representing a range of rows and returns a slice object.

    Args:
        range (str): A string representing the range of rows in the format 'start:end'.
                     Both 'start' and 'end' are 1-based indices. If 'start' is omitted,
                     it defaults to the beginning. If 'end' is omitted, it defaults to
                     the end.

    Returns:
        slice: A slice object representing the range of rows. The start index is 0-based,
               and the end index is exclusive.

    """
    start, end = range.split(RANGE_SPLITTER)
    start_index = int(start) - 1 if start else 0
    end_index = int(end) if end else None
    return slice(start_index, end_index)


def parse_cells(mapping: dict, sheet: DataFrame) -> dict:
    """
    Parses specific cells from a DataFrame based on a given cell map.
    
    Args:
        mapping (dict): A dictionary where keys are cell names and values are cell locations in Excel notation (e.g., 'A1', 'B2').
        sheet (DataFrame): A pandas DataFrame representing the Excel sheet to parse cells from.
   
    Returns:
        dict: A dictionary where keys are cell names and values are the corresponding cell values from the DataFrame.
    
    """
    result = {}
    for name, location in mapping.items():
        result[name] = sheet.iloc[xl_cell_to_rowcol(location)]
    
    return result


def parse_rows(mapping: dict, sheet: DataFrame) -> dict:
    """
    Parses specified rows from a DataFrame based on a given row map.
    
    Args:
        mapping (dict): A dictionary where keys are row names and values are row ranges in a format that can be parsed by `parse_range`.
        sheet (DataFrame): The DataFrame from which rows will be extracted.
    
    Returns:
        dict: A dictionary where keys are row names and values are lists of row values.
        
    """
    result = {}
    for name, range in mapping.items():
        row = sheet.iloc[parse_range(range)]
        result[name] = row.values.flatten().tolist()
    
    return result



def parse_repeat_row(mapping: dict, sheet: DataFrame) -> list:
    """
    Parses rows from a DataFrame based on a mapping configuration.

    Args:
        mapping (dict): A dictionary containing the mapping configuration. 
                       It should have the keys:
                       - 'range': A string representing the range of rows to parse.
                       - 'components': A dictionary where keys are the names of the 
                                       components and values are the column ranges.
        sheet (DataFrame): The DataFrame to parse the rows from.
    
    Returns:
        list: A list of dictionaries, where each dictionary represents a parsed row 
              with keys as component names and values as the corresponding cell values 
              or lists of cell values.

    """
    result = []
    for row_index in sheet.index[parse_row_range(mapping['range'])]:
        row = {}
        for name, column in mapping["components"].items():
            _, col_indices = parse_range(column)
            if isinstance(col_indices, int):
                row[name] = sheet.iloc[row_index, col_indices]
            else:
                part = sheet.iloc[[row_index], col_indices]
                row[name] = part.values.flatten().tolist()
        result.append(row)
    
    return result



def parse_repeat_rows(mapping: list, sheet: DataFrame) -> dict:
    """
    Parses repeated rows in a given DataFrame based on a mapping.
    
    Args:
        mapping (list): A list of dictionaries where each dictionary contains 
                        the 'name' key and other necessary information for parsing.
        sheet (DataFrame): The DataFrame containing the data to be parsed.
    
    Returns:
        dict: A dictionary where the keys are the 'name' values from the mapping 
              and the values are the results of the parse_repeat_row function.
    """
    result = {}
    for item in mapping:
        result[item['name']] = parse_repeat_row(item, sheet)
    
    return result
