import json
from pandas import read_excel
from excel_parser import parse_cells, parse_rows, parse_repeat_rows


def load_json() -> dict:
    """
    Load and return the contents of a JSON file.

    This function opens the file 'v7.json' in read mode with UTF-8 encoding,
    reads its contents, and returns the parsed JSON data as a Python object.

    Returns:
        dict: The parsed JSON data from the file.

    Raises:
        FileNotFoundError: If the file 'v7.json' does not exist.
        json.JSONDecodeError: If the file contains invalid JSON.

    """
    with open('v7.json', mode='r', encoding='utf-8') as f:
        return json.load(f)


def parse_excel_v7(excel: bytes) -> dict:
    """
    Parses an Excel file and extracts data based on a predefined mapping.
    
    Args:
        excel (bytes): The Excel file content in bytes.
    
    Returns:
        dict: A dictionary containing parsed data with the following keys:
            - 'cells': Parsed cell data.
            - 'rows': Parsed row data.
            - 'repeatRows': Parsed repeat row data.

    """
    mapping = load_json()
    sheet = read_excel(excel, header=None, keep_default_na=False)

    result = {}
    result['cells'] = parse_cells(mapping['cells'], sheet)
    result['rows'] = parse_rows(mapping['rows'], sheet)
    result['repeatRows'] = parse_repeat_rows(mapping['repeatRows'], sheet)

    return result


if __name__ == '__main__':
    with open('report.xlsm', 'rb') as f:
        print(parse_excel_v7(f.read()))
