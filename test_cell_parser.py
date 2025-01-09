import unittest
import pandas as pd
from excel_parser import parse_cells

class TestParseCells(unittest.TestCase):

    def setUp(self):
        # Create a sample DataFrame to use in tests
        data = {
            'A': [1, 2, 3],
            'B': [4, 5, 6],
            'C': [7, 8, 9]
        }
        self.sheet = pd.DataFrame(data)

    def test_parse_single_cell(self):
        cell_map = {'cell1': 'A1'}
        result = parse_cells(cell_map, self.sheet)
        expected = {'cell1': 1}
        self.assertEqual(result, expected)

    def test_parse_multiple_cells(self):
        cell_map = {'cell1': 'A1', 'cell2': 'B2', 'cell3': 'C3'}
        result = parse_cells(cell_map, self.sheet)
        expected = {'cell1': 1, 'cell2': 5, 'cell3': 9}
        self.assertEqual(result, expected)

    def test_parse_nonexistent_cell(self):
        cell_map = {'cell1': 'D4'}
        with self.assertRaises(IndexError):
            parse_cells(cell_map, self.sheet)

if __name__ == '__main__':
    unittest.main()