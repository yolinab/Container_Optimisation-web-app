"""
Tests for app/utils/validate.py — specifically the end-to-end pallet count
reconciliation check added after a real-world silent data loss bug (a
Kim Phat order file showed 488 pallets in the source sheet but only 463
made it into the packing report, with zero errors or warnings raised).
"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "app"))

from utils.validate import validate_packing_result


L_CM     = 1203
HDOOR_CM = 259
WMAX_KG  = 18000
GAP_CM   = 5


class _FakeBlock:
    """Duck-typed stand-in for BlockInstance — validate.py only reads .block_id/.value."""
    def __init__(self, block_id, value):
        self.block_id = block_id
        self.value = value


def _container(idx, block_id, pallet_count, used_length_cm=115):
    return {
        "container_index": idx,
        "used_length_cm": used_length_cm,
        "leftover_cm": L_CM - used_length_cm,
        "loaded_weight": 0,
        "rows": [{
            "block_id": block_id,
            "pallet_count": pallet_count,
            "height_cm": 200,
            "length_cm": used_length_cm,
            "y_start_cm": 0,
        }],
    }


def _issues_by_code(issues, code):
    return [i for i in issues if i["code"] == code]


class TestPalletReconciliation:

    def test_passes_when_fully_accounted(self):
        """packed + parse_dropped + build_dropped == raw_input -> no error."""
        blocks = [_FakeBlock(1, 8)]
        containers = [_container(1, block_id=1, pallet_count=8)]
        issues = validate_packing_result(
            containers=containers, original_blocks=blocks, np_boxes=[],
            L_cm=L_CM, Hdoor_cm=HDOOR_CM, Wmax_kg=WMAX_KG, gap_cm=GAP_CM,
            raw_input_pallets=8, parse_dropped_pallets=0, build_dropped_pallets=0,
        )
        assert _issues_by_code(issues, "TOTAL_PALLET_COUNT_MISMATCH") == []

    def test_passes_when_gap_fully_explained_by_dropped_counts(self):
        """13 ordered, 5 dropped during parsing (warned about), 8 packed -> reconciles cleanly."""
        blocks = [_FakeBlock(1, 8)]
        containers = [_container(1, block_id=1, pallet_count=8)]
        issues = validate_packing_result(
            containers=containers, original_blocks=blocks, np_boxes=[],
            L_cm=L_CM, Hdoor_cm=HDOOR_CM, Wmax_kg=WMAX_KG, gap_cm=GAP_CM,
            raw_input_pallets=13, parse_dropped_pallets=5, build_dropped_pallets=0,
        )
        assert _issues_by_code(issues, "TOTAL_PALLET_COUNT_MISMATCH") == []

    def test_fires_on_unexplained_gap(self):
        """
        This is the exact real-world bug: 25 pallets missing with no dropped
        count and no warning explaining them. Must raise a hard ERROR.
        """
        blocks = [_FakeBlock(1, 463)]
        containers = [_container(1, block_id=1, pallet_count=463)]
        issues = validate_packing_result(
            containers=containers, original_blocks=blocks, np_boxes=[],
            L_cm=L_CM, Hdoor_cm=HDOOR_CM, Wmax_kg=WMAX_KG, gap_cm=GAP_CM,
            raw_input_pallets=488, parse_dropped_pallets=0, build_dropped_pallets=0,
        )
        mismatches = _issues_by_code(issues, "TOTAL_PALLET_COUNT_MISMATCH")
        assert len(mismatches) == 1
        assert mismatches[0]["level"] == "ERROR"
        assert "488" in mismatches[0]["message"]
        assert "463" in mismatches[0]["message"]
        assert "25" in mismatches[0]["message"]

    def test_skipped_when_raw_input_pallets_not_provided(self):
        """Backward-compat: omitting raw_input_pallets must not raise or crash."""
        blocks = [_FakeBlock(1, 8)]
        containers = [_container(1, block_id=1, pallet_count=8)]
        issues = validate_packing_result(
            containers=containers, original_blocks=blocks, np_boxes=[],
            L_cm=L_CM, Hdoor_cm=HDOOR_CM, Wmax_kg=WMAX_KG, gap_cm=GAP_CM,
        )
        assert _issues_by_code(issues, "TOTAL_PALLET_COUNT_MISMATCH") == []

    def test_row_skip_warnings_surfaced_as_issues(self):
        """Row-level skip explanations must show up as WARNING-level issues."""
        blocks = [_FakeBlock(1, 8)]
        containers = [_container(1, block_id=1, pallet_count=8)]
        issues = validate_packing_result(
            containers=containers, original_blocks=blocks, np_boxes=[],
            L_cm=L_CM, Hdoor_cm=HDOOR_CM, Wmax_kg=WMAX_KG, gap_cm=GAP_CM,
            raw_input_pallets=13, parse_dropped_pallets=5, build_dropped_pallets=0,
            row_skip_warnings=["Row 12: dimension 'xyz' could not be parsed — 5 pallet(s) skipped."],
        )
        skipped = _issues_by_code(issues, "ROW_SKIPPED")
        assert len(skipped) == 1
        assert skipped[0]["level"] == "WARNING"
        assert "Row 12" in skipped[0]["message"]

    def test_gap_larger_than_explained_still_fires(self):
        """A partially-explained gap (some dropped, but not enough) must still error."""
        blocks = [_FakeBlock(1, 8)]
        containers = [_container(1, block_id=1, pallet_count=8)]
        issues = validate_packing_result(
            containers=containers, original_blocks=blocks, np_boxes=[],
            L_cm=L_CM, Hdoor_cm=HDOOR_CM, Wmax_kg=WMAX_KG, gap_cm=GAP_CM,
            raw_input_pallets=20, parse_dropped_pallets=5, build_dropped_pallets=0,
        )
        # 8 packed + 5 dropped = 13, but raw was 20 -> 7 pallets still unexplained
        mismatches = _issues_by_code(issues, "TOTAL_PALLET_COUNT_MISMATCH")
        assert len(mismatches) == 1
        assert "7" in mismatches[0]["message"]
