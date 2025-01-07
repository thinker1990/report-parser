import unittest
from parse_excel import range_slice
from parse_excel import parse_row_range
from parse_excel import parse_length

class TestParseExcel(unittest.TestCase):
    
    def test_range_slice_valid(self):
        self.assertEqual(range_slice("A1:B2"), (slice(0, 2), slice(0, 2)))
        self.assertEqual(range_slice("C3:D4"), (slice(2, 4), slice(2, 4)))
        self.assertEqual(range_slice("E5:F6"), (slice(4, 6), slice(4, 6)))
    
    def test_range_slice_single_cell(self):
        self.assertEqual(range_slice("A1:A1"), (slice(0, 1), slice(0, 1)))
        self.assertEqual(range_slice("B2:B2"), (slice(1, 2), slice(1, 2)))
    
    def test_range_slice_invalid_format(self):
        with self.assertRaises(ValueError):
            range_slice("A1B2")
        with self.assertRaises(ValueError):
            range_slice("A1:B2:C3")
    
    def test_parse_row_range_valid(self):
        self.assertEqual(parse_row_range("1:5"), slice(0, 5))
        self.assertEqual(parse_row_range("2:10"), slice(1, 10))
        self.assertEqual(parse_row_range("3:"), slice(2, None))
        self.assertEqual(parse_row_range(":5"), slice(0, 5))
        self.assertEqual(parse_row_range(":"), slice(0, None))
    
    def test_parse_row_range_invalid_format(self):
        with self.assertRaises(ValueError):
            parse_row_range("1-5")
        with self.assertRaises(ValueError):
            parse_row_range("1:5:10")

    def test_parse_length_with_int(self):
        self.assertEqual(parse_length(5, {}), 5)
        self.assertEqual(parse_length(0, {}), 0)
        self.assertEqual(parse_length(-3, {}), -3)
   
    def test_parse_length_with_str(self):
        lookup = {'length1': 10, 'length2': 20}
        self.assertEqual(parse_length('@length1', lookup), 10)
        self.assertEqual(parse_length('@length2', lookup), 20)
   
    def test_parse_length_with_str_no_at(self):
        lookup = {'length1': 10, 'length2': 20}
        self.assertEqual(parse_length('length1', lookup), 10)
        self.assertEqual(parse_length('length2', lookup), 20)
   
    def test_parse_length_with_invalid_str(self):
        lookup = {'length1': 10, 'length2': 20}
        with self.assertRaises(KeyError):
            parse_length('@length3', lookup)
        with self.assertRaises(KeyError):
            parse_length('length3', lookup)



if __name__ == '__main__':
    unittest.main()
