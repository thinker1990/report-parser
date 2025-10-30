import unittest
import pandas as pd
from excel_parser import (
    extract_cell_value, 
    extract_range_values,
    parse_cells,
    parse_ranges,
    parse_row_components,
    parse_repeat_rows
)


class TestExtractHelpers(unittest.TestCase):
    """Test the new helper functions that reduce code duplication."""

    def setUp(self):
        # Create a sample DataFrame to use in tests
        data = {
            'A': [1, 2, 3, 4],
            'B': [5, 6, 7, 8],
            'C': [9, 10, 11, 12],
            'D': [13, 14, 15, 16]
        }
        self.sheet = pd.DataFrame(data)

    def test_extract_cell_value(self):
        """Test extracting a single cell value."""
        result = extract_cell_value('A1', self.sheet)
        self.assertEqual(result, 1)
        
        result = extract_cell_value('C3', self.sheet)
        self.assertEqual(result, 11)

    def test_extract_cell_value_with_row_override(self):
        """Test extracting a cell value with row override."""
        # Location says A1 (row 0), but we want row 2
        result = extract_cell_value('A1', self.sheet, row_index=2)
        self.assertEqual(result, 3)  # Value at A3
        
        # Location says B2 (row 1), but we want row 3
        result = extract_cell_value('B2', self.sheet, row_index=3)
        self.assertEqual(result, 8)  # Value at B4

    def test_extract_range_values(self):
        """Test extracting values from a range."""
        result = extract_range_values('A1:B2', self.sheet)
        expected = [1, 5, 2, 6]
        self.assertEqual(result, expected)

    def test_extract_range_values_single_row(self):
        """Test extracting values from a range in a specific row."""
        result = extract_range_values('A1:C1', self.sheet, row_index=1)
        expected = [2, 6, 10]
        self.assertEqual(result, expected)

    def test_parse_cells(self):
        """Test parsing multiple cells."""
        cell_map = {'cell1': 'A1', 'cell2': 'B2', 'cell3': 'C3'}
        result = parse_cells(cell_map, self.sheet)
        expected = {'cell1': 1, 'cell2': 6, 'cell3': 11}
        self.assertEqual(result, expected)

    def test_parse_ranges(self):
        """Test parsing multiple ranges."""
        range_map = {
            'range1': 'A1:B1',
            'range2': 'C2:D2'
        }
        result = parse_ranges(range_map, self.sheet)
        expected = {
            'range1': [1, 5],
            'range2': [10, 14]
        }
        self.assertEqual(result, expected)

    def test_parse_row_components_cells_only(self):
        """Test parsing row components with single cells."""
        mapping = {
            'col1': 'A1',
            'col2': 'B1',
            'col3': 'C1'
        }
        result = parse_row_components(1, mapping, self.sheet)
        expected = {'col1': 2, 'col2': 6, 'col3': 10}
        self.assertEqual(result, expected)

    def test_parse_row_components_with_range(self):
        """Test parsing row components with both cells and ranges."""
        mapping = {
            'single': 'A1',
            'range': 'B1:D1'
        }
        result = parse_row_components(2, mapping, self.sheet)
        expected = {
            'single': 3,
            'range': [7, 11, 15]
        }
        self.assertEqual(result, expected)


class TestParseRepeatRows(unittest.TestCase):
    """Test the parse_repeat_rows function."""

    def setUp(self):
        # Create a larger sample DataFrame
        data = {
            'A': [1, 2, 3, 4, 5],
            'B': [6, 7, 8, 9, 10],
            'C': [11, 12, 13, 14, 15]
        }
        self.sheet = pd.DataFrame(data)

    def test_parse_repeat_rows(self):
        """Test parsing repeat rows with a range."""
        mapping = {
            'range': '2:4',
            'components': {
                'colA': 'A1',
                'colB': 'B1'
            }
        }
        result = parse_repeat_rows(mapping, self.sheet)
        expected = [
            {'colA': 2, 'colB': 7},
            {'colA': 3, 'colB': 8},
            {'colA': 4, 'colB': 9}
        ]
        self.assertEqual(result, expected)

    def test_parse_repeat_rows_open_ended(self):
        """Test parsing repeat rows with an open-ended range."""
        mapping = {
            'range': '3:',
            'components': {
                'value': 'C1'
            }
        }
        result = parse_repeat_rows(mapping, self.sheet)
        expected = [
            {'value': 13},
            {'value': 14},
            {'value': 15}
        ]
        self.assertEqual(result, expected)


if __name__ == '__main__':
    unittest.main()
