import unittest
from datetime import datetime

from sheets import format_date_label, tab_name, _parse_amount


class SheetHelperTests(unittest.TestCase):
    def test_format_date_label(self) -> None:
        dt = datetime(2026, 9, 2)
        self.assertEqual(format_date_label(dt), "Sep 2")

    def test_tab_name(self) -> None:
        dt = datetime(2026, 9, 2)
        self.assertEqual(tab_name(dt, "Goober", 123), "September - Goober")

    def test_parse_amount(self) -> None:
        self.assertEqual(_parse_amount("25.50"), 25.5)
        self.assertEqual(_parse_amount(""), 0.0)
        self.assertEqual(_parse_amount("$1,234.50"), 1234.5)


if __name__ == "__main__":
    unittest.main()
