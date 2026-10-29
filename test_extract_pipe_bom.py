import unittest
from pathlib import Path
from unittest.mock import patch

import extract_pipe_bom


class ExtractFittingRowsTests(unittest.TestCase):
    def test_removes_duplicate_item_number_from_description(self) -> None:
        values = [
            "FITTINGS",
            "6",
            "6TEE B16.9 A420 WPL6-SMLS",
            "4 x 4",
            "I63993118",
            "1",
            "7",
            "6 x 4 CONCENTRIC REDUCER B16.9 A420 WPL6-SMLS",
            "6 x 4",
            "I63992978",
            "1",
            "FLANGES",
        ]
        entities = [
            extract_pipe_bom.TextEntity(text, "GT_1") for text in values
        ]

        with patch.object(
            extract_pipe_bom, "extract_modelspace_text_entities", return_value=entities
        ):
            rows = list(
                extract_pipe_bom.extract_fitting_rows(Path("sample_Rev0.dxf"))
            )

        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0][5], "TEE B16.9 A420 WPL6-SMLS")
        self.assertEqual(rows[1][5], "6 x 4 CONCENTRIC REDUCER B16.9 A420 WPL6-SMLS")


if __name__ == "__main__":
    unittest.main()
