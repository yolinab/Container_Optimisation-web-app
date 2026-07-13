"""
Tests for app/utils/parse_xlsx.py
"""
import pytest
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "app"))

from utils.parse_xlsx import (
    parse_pallet_excel_v3,
    parse_np_boxes_excel_v3,
    _parse_pallet_size_str,
    _detect_header_row,
)


# ---------------------------------------------------------------------------
# _parse_pallet_size_str
# ---------------------------------------------------------------------------

class TestParsePalletSizeStr:
    def test_comma_decimal(self):
        assert _parse_pallet_size_str("1,15x1,15x1,20") == (115, 115, 120)

    def test_dot_decimal(self):
        assert _parse_pallet_size_str("1.15x1.15x1.20") == (115, 115, 120)

    def test_with_cm_suffix(self):
        # Some exports add 'cm' text
        assert _parse_pallet_size_str("1,15x0,77x1,27cm") == (115, 77, 127)

    def test_mixed_separators(self):
        assert _parse_pallet_size_str("1.15x0.77x1.27") == (115, 77, 127)

    def test_rounding(self):
        # 1.149 * 100 = 114.9 → rounds to 115
        assert _parse_pallet_size_str("1.149x0.77x1.27") == (115, 77, 127)

    def test_wrong_parts_raises(self):
        with pytest.raises(ValueError, match="Cannot parse"):
            _parse_pallet_size_str("1.15x1.15")  # only 2 parts

    def test_non_numeric_raises(self):
        with pytest.raises((ValueError, Exception)):
            _parse_pallet_size_str("axbxc")

    def test_np_box_small_dims(self):
        # 0.46x0.46x0.87 → 46x46x87
        assert _parse_pallet_size_str("0,46x0,46x0,87") == (46, 46, 87)


# ---------------------------------------------------------------------------
# _detect_header_row
# ---------------------------------------------------------------------------

class TestDetectHeaderRow:
    def test_finds_header_at_row_0(self, excel_minimal_valid):
        row = _detect_header_row(excel_minimal_valid)
        assert row == 0

    def test_finds_header_after_blank_rows(self, excel_blank_header_rows):
        row = _detect_header_row(excel_blank_header_rows)
        assert row == 2  # two blank rows before actual header


# ---------------------------------------------------------------------------
# parse_pallet_excel_v3
# ---------------------------------------------------------------------------

class TestParsePalletExcelV3:
    def test_minimal_valid_count(self, excel_minimal_valid):
        lengths, widths, heights, pallets_data, meta, _ = parse_pallet_excel_v3(excel_minimal_valid)
        assert len(meta) == 8                   # 8 physical pallets
        assert len(pallets_data) == 1           # 1 type row

    def test_minimal_valid_dimensions(self, excel_minimal_valid):
        lengths, widths, heights, _, meta, _ = parse_pallet_excel_v3(excel_minimal_valid)
        assert all(l == 115 for l in lengths)
        assert all(w == 115 for w in widths)
        assert all(h == 120 for h in heights)

    def test_np_rows_excluded(self, excel_mixed):
        """NP rows must NOT appear in pallet parse results."""
        _, _, _, pallets_data, meta, _ = parse_pallet_excel_v3(excel_mixed)
        for p in meta:
            assert p["height"] != 87   # NP box height is 87 cm

    def test_mixed_counts(self, excel_mixed):
        """A2 (8) + C2 (4) = 12 physical pallets."""
        _, _, _, _, meta, _ = parse_pallet_excel_v3(excel_mixed)
        assert len(meta) == 12

    def test_dot_decimal_dimensions(self, excel_dot_decimal):
        _, _, heights, _, meta, _ = parse_pallet_excel_v3(excel_dot_decimal)
        assert all(h == 120 for h in heights)

    def test_blank_header_rows(self, excel_blank_header_rows):
        _, _, _, _, meta, _ = parse_pallet_excel_v3(excel_blank_header_rows)
        assert len(meta) == 8

    def test_missing_required_column_raises(self, excel_wrong_col_name):
        """When count column not found, KeyError should contain helpful info."""
        with pytest.raises(KeyError) as exc_info:
            parse_pallet_excel_v3(excel_wrong_col_name)
        msg = str(exc_info.value)
        # Should mention what columns were tried or what was found
        assert "Columns" in msg or "column" in msg.lower() or "found" in msg.lower()

    def test_count_col_override_success(self, excel_wrong_col_override):
        """count_col_override='MyQty' should work even though default fuzzy match fails."""
        _, _, _, _, meta, _ = parse_pallet_excel_v3(excel_wrong_col_override, count_col_override="MyQty")
        assert len(meta) == 8

    def test_count_col_override_not_in_file_raises(self, excel_minimal_valid):
        """If override column doesn't exist, should raise with clear message."""
        with pytest.raises(KeyError) as exc_info:
            parse_pallet_excel_v3(excel_minimal_valid, count_col_override="NonExistentColumn")
        assert "NonExistentColumn" in str(exc_info.value)

    def test_meta_has_required_keys(self, excel_minimal_valid):
        _, _, _, _, meta, _ = parse_pallet_excel_v3(excel_minimal_valid)
        required = {"pallet_id", "length", "width", "height"}
        for m in meta:
            assert required.issubset(m.keys())

    def test_np_only_returns_empty(self, excel_np_only):
        """NP-only file should return 0 pallets."""
        _, _, _, pallets_data, meta, _ = parse_pallet_excel_v3(excel_np_only)
        assert len(meta) == 0

    def test_raw_total_qty_matches_valid_rows(self, excel_minimal_valid):
        """No dropped rows -> raw_total_qty equals the actual pallet count."""
        _, _, _, _, meta, diag = parse_pallet_excel_v3(excel_minimal_valid)
        assert diag["raw_total_qty"] == 8
        assert diag["raw_total_qty"] == len(meta)
        assert diag["warnings"] == []

    def test_bad_dimension_row_tracked_not_silently_dropped(self, excel_bad_dimension_row):
        """
        A row with a real quantity but a garbled dimension string must be
        excluded from meta_per_pallet AND explained via a warning AND still
        counted in raw_total_qty — this is the exact bug that caused a
        real-world 25-pallet silent gap.
        """
        _, _, _, _, meta, diag = parse_pallet_excel_v3(excel_bad_dimension_row)
        assert len(meta) == 8                      # only the good row parsed
        assert diag["raw_total_qty"] == 13          # 8 good + 5 from the bad row
        assert len(diag["warnings"]) == 1
        assert "5 pallet(s) skipped" in diag["warnings"][0]
        # The gap must be fully explained: raw total - parsed = the warned-about count
        assert diag["raw_total_qty"] - len(meta) == 5

    def test_blank_quantity_row_excluded_without_warning(self, excel_blank_qty_row):
        """A row with a blank/zero quantity is a normal 'not ordered' case —
        must be excluded silently, with no warning and no raw_total_qty impact."""
        _, _, _, _, meta, diag = parse_pallet_excel_v3(excel_blank_qty_row)
        assert len(meta) == 8
        assert diag["raw_total_qty"] == 8
        assert diag["warnings"] == []


# ---------------------------------------------------------------------------
# parse_np_boxes_excel_v3
# ---------------------------------------------------------------------------

class TestParseNpBoxesV3:
    def test_np_only_finds_boxes(self, excel_np_only):
        boxes, _ = parse_np_boxes_excel_v3(excel_np_only)
        assert len(boxes) == 1
        assert boxes[0]["quantity"] == 10

    def test_np_boxes_in_mixed(self, excel_mixed):
        boxes, _ = parse_np_boxes_excel_v3(excel_mixed)
        assert len(boxes) == 1
        assert boxes[0]["quantity"] == 6

    def test_no_np_in_pallet_only(self, excel_minimal_valid):
        boxes, _ = parse_np_boxes_excel_v3(excel_minimal_valid)
        assert boxes == []

    def test_np_box_dimensions_cm(self, excel_np_only):
        boxes, _ = parse_np_boxes_excel_v3(excel_np_only)
        b = boxes[0]
        assert b["length_cm"] == 46
        assert b["width_cm"] == 46
        assert b["height_cm"] == 87

    def test_np_box_has_volume(self, excel_np_only):
        boxes, _ = parse_np_boxes_excel_v3(excel_np_only)
        b = boxes[0]
        assert b["volume_cm3"] == 46 * 46 * 87
        assert b["total_volume_cm3"] == b["volume_cm3"] * 10
