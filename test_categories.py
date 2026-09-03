import difflib
import unittest

from categories import CategoryMatch


class CategoryMatchLogicTests(unittest.TestCase):
    def test_fuzzy_match(self) -> None:
        known = ["groceries", "rent", "coffee"]
        suggestions = difflib.get_close_matches(
            "grocieries", known, n=1, cutoff=0.75
        )
        self.assertEqual(suggestions, ["groceries"])
        match = CategoryMatch(kind="fuzzy", suggested=suggestions[0])
        self.assertEqual(match.suggested, "groceries")


if __name__ == "__main__":
    unittest.main()
