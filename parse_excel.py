import json
from pandas import read_excel
from xlsxwriter.utility import xl_cell_to_rowcol

def range_slice(range_str):
    """
    将表示Excel单元格范围的字符串转换为两个slice对象，分别表示行和列的切片范围。

    Args:
        range_str (str): 表示Excel单元格范围的字符串，格式为"A1:B2"。

    Returns:
        Tuple[slice, slice]: 包含两个slice对象的元组，分别表示行和列的切片范围。

    """
    (start_row, start_col), (stop_row, stop_col) = [
        xl_cell_to_rowcol(c) for c in range_str.split(':')]
    return slice(start_row, stop_row + 1), slice(start_col, stop_col + 1)

def parse_row_range(range_str):
    """
    解析行范围字符串并返回对应的slice对象。

    参数:
        range_str (str): 行范围字符串，格式为"start:end"，其中start和end可以是数字或空字符串。

    返回:
        slice: 解析得到的slice对象。

    抛出:
        ValueError: 如果range_str的格式不正确。
    """
    try:
        start, end = range_str.split(':')
    except ValueError:
        raise ValueError("range_str的格式不正确，应为'start:end'。")

    start = int(start) - 1 if start else 0
    end = int(end) if end else None
    return slice(start, end, None)  # 显式设置step为None

def parse_length(length, lookup):
    """
    根据传入的长度值返回相应的长度。

    Args:
        length (int|str): 传入的长度值，可以是整数类型或字符串类型。
        lookup (dict): 一个字典，用于将字符串类型的长度值映射到对应的整数长度值。

    Returns:
        int: 返回解析后的长度值，如果传入的length是整数类型，则直接返回该整数；
             如果传入的length是字符串类型，则从lookup中查找对应的整数长度值并返回。

    Raises:
        KeyError: 如果传入的length是字符串类型，并且在lookup中找不到对应的键，则会引发KeyError异常。
    """
    if isinstance(length, int):
        return length
    if isinstance(length, str):
        return lookup[length.strip('@')]


def load_json():
    """
    从文件中加载JSON数据。

    Args:
        无参数。

    Returns:
        dict: 从文件中加载的JSON数据。
    """
    with open('v7.json', mode='r', encoding='utf-8') as f:
        return json.load(f)

def parse_excel_v7(excel_content):
    """
    从Excel文件中解析数据，并返回字典格式的解析结果。

    Args:
        无

    Returns:
        dict: 包含解析结果的字典。

    """
    result = {}
    wb = read_excel(excel_content, header=None, keep_default_na=False)

    js = load_json()
    for i in js["cell"]:
        result[i['name']] = wb.iloc[xl_cell_to_rowcol(i['location'])]

    for i in js["row"]:
        row = wb.iloc[range_slice(i['range'])].values.flatten()
        result[i['name']] = row[:parse_length(i['length'], result)].tolist()

    for i in js["area"]:
        result['area'] = []
        rows = wb.index[parse_row_range(i['range'])]
        for idx in rows:
            fai = {}
            for j in i["components"]:
                col = j['column']
                if ':' in col:
                    _, col_slice = range_slice(col)
                    part = wb.iloc[[idx], col_slice].values.flatten()
                    fai[j['name']] = part[:parse_length(j['length'], result)].tolist()
                else:
                    col_index = xl_cell_to_rowcol(col)[1]
                    fai[j['name']] = wb.iloc[idx, col_index]
            result['area'].append(fai)

    return result
