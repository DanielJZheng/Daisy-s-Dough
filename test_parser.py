import unittest

from parser import parse_message


class ParseMessageTests(unittest.TestCase):
    def test_expense_basic(self) -> None:
        entry = parse_message("25.50 groceries")
        assert entry is not None
        self.assertEqual(entry.entry_type, "expense")
        self.assertAlmostEqual(entry.amount, 25.50)
        self.assertEqual(entry.category, "groceries")
        self.assertIsNone(entry.note)

    def test_expense_with_note(self) -> None:
        entry = parse_message("25.50 groceries lunch at Costco")
        assert entry is not None
        self.assertEqual(entry.note, "lunch at Costco")

    def test_income_basic(self) -> None:
        entry = parse_message("income 2500 salary")
        assert entry is not None
        self.assertEqual(entry.entry_type, "income")
        self.assertAlmostEqual(entry.amount, 2500.0)
        self.assertEqual(entry.category, "salary")

    def test_income_with_note(self) -> None:
        entry = parse_message("income 500 freelance side gig")
        assert entry is not None
        self.assertEqual(entry.note, "side gig")

    def test_invalid(self) -> None:
        self.assertIsNone(parse_message("hello world"))
        self.assertIsNone(parse_message("income salary"))


if __name__ == "__main__":
    unittest.main()
