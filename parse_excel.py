import json
from pandas import read_excel
from excel_parser import parse_cells, parse_ranges, parse_repeat_rows
from io import BytesIO


def load_json() -> dict:
    with open("v7.json", mode="r", encoding="utf-8") as f:
        return json.load(f)


def parse_excel_v7(excel: bytes) -> dict:
    mapping = load_json()
    sheet = read_excel(BytesIO(excel), header=None, keep_default_na=False)

    result = {}
    result["cells"] = parse_cells(mapping["cells"], sheet)
    result["ranges"] = parse_ranges(mapping["ranges"], sheet)
    result["repeatRows"] = parse_repeat_rows(mapping["repeatRows"], sheet)

    return result


if __name__ == "__main__":
    with open("report.xlsm", "rb") as f:
        print(parse_excel_v7(f.read()))
