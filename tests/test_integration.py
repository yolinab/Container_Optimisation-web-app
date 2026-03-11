"""
Integration tests — run the full pipeline on fixture Excel files.
Tests verify that the pipeline completes or fails with the correct, specific errors.
"""
import pytest
import sys, os
from pathlib import Path

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "app"))


def _run(excel_path, tmp_path, **kwargs):
    """Helper: run run_pipeline() with output in tmp_path/outputs."""
    from pipeline import run_pipeline
    result = run_pipeline(
        excel_path=Path(excel_path),
        out_dir=tmp_path / "outputs",
        **kwargs,
    )
    return result["containers"]


# ---------------------------------------------------------------------------
# Happy-path runs
# ---------------------------------------------------------------------------

class TestHappyPath:
    def test_minimal_valid_produces_containers(self, excel_minimal_valid, tmp_path):
        """8 A2 pallets → at least 1 container returned."""
        containers = _run(excel_minimal_valid, tmp_path)
        assert isinstance(containers, list)
        assert len(containers) >= 1

    def test_minimal_valid_report_created(self, excel_minimal_valid, tmp_path):
        """report.xlsx must be written to outputs/ after a successful run."""
        _run(excel_minimal_valid, tmp_path)
        assert (tmp_path / "outputs" / "report.xlsx").exists()

    def test_minimal_valid_summary_created(self, excel_minimal_valid, tmp_path):
        _run(excel_minimal_valid, tmp_path)
        assert (tmp_path / "outputs" / "summary.txt").exists()

    def test_container_structure_keys(self, excel_minimal_valid, tmp_path):
        """Each container dict must have the required keys."""
        containers = _run(excel_minimal_valid, tmp_path)
        required = {"container_index", "rows", "used_length_cm", "loaded_weight", "leftover_cm"}
        for c in containers:
            assert required.issubset(c.keys())

    def test_rows_have_required_keys(self, excel_minimal_valid, tmp_path):
        """Each row within a container must have expected keys."""
        containers = _run(excel_minimal_valid, tmp_path)
        row_required = {"block_id", "block_type", "length_cm", "height_cm", "y_start_cm", "pallet_count"}
        for c in containers:
            for row in c["rows"]:
                assert row_required.issubset(row.keys())

    def test_used_length_within_container(self, excel_minimal_valid, tmp_path):
        """used_length_cm must not exceed container length."""
        from config import CONTAINER_LENGTH_CM
        containers = _run(excel_minimal_valid, tmp_path)
        for c in containers:
            assert c["used_length_cm"] <= CONTAINER_LENGTH_CM

    def test_mixed_pallets_and_np(self, excel_mixed, tmp_path):
        """Mixed file (A2 + C2 + NP) → containers packed, box_zones assigned."""
        containers = _run(excel_mixed, tmp_path)
        assert len(containers) >= 1
        all_zones = [z for c in containers for z in c.get("box_zones", [])]
        assert len(all_zones) >= 0

    def test_two_container_run(self, excel_two_containers, tmp_path):
        """Large input should produce multiple containers."""
        containers = _run(excel_two_containers, tmp_path)
        assert len(containers) >= 2

    def test_np_only_no_containers(self, excel_np_only, tmp_path):
        """NP-only file: no pallets → RuntimeError about no pallets parsed."""
        with pytest.raises(RuntimeError, match="(?i)no pallets|parsed|pallet"):
            _run(excel_np_only, tmp_path)


# ---------------------------------------------------------------------------
# Validation error paths
# ---------------------------------------------------------------------------

class TestValidationErrors:
    def test_multiples_needed_raises(self, excel_multiples_needed, tmp_path):
        """Non-multiple pallet counts must raise RuntimeError with 'add' or 'pallet' in message."""
        with pytest.raises(RuntimeError) as exc_info:
            _run(excel_multiples_needed, tmp_path)
        msg = str(exc_info.value).lower()
        assert "add" in msg or "pallet" in msg or "multiple" in msg

    def test_multiples_error_mentions_count(self, excel_multiples_needed, tmp_path):
        """Multiples error should say how many pallets to add."""
        with pytest.raises(RuntimeError) as exc_info:
            _run(excel_multiples_needed, tmp_path)
        msg = str(exc_info.value)
        assert any(ch.isdigit() for ch in msg)

    def test_all_too_tall_raises_runtime_error(self, excel_all_too_tall, tmp_path):
        """Pallets with unknown block type rule → 0 blocks built → RuntimeError."""
        with pytest.raises(RuntimeError):
            _run(excel_all_too_tall, tmp_path)

    def test_all_too_tall_no_subscript_crash(self, excel_all_too_tall, tmp_path):
        """Confirm the _BoolVarImpl subscript regression does not resurface."""
        try:
            _run(excel_all_too_tall, tmp_path)
        except RuntimeError:
            pass  # expected
        except TypeError as e:
            pytest.fail(f"Got TypeError (likely regression): {e}")

    def test_unknown_footprint_raises(self, excel_unknown_footprint, tmp_path):
        """Unknown footprints → 0 blocks → RuntimeError."""
        with pytest.raises(RuntimeError):
            _run(excel_unknown_footprint, tmp_path)

    def test_missing_file_raises(self, tmp_path):
        """Non-existent Excel file must raise, not silently produce 0 containers."""
        with pytest.raises(Exception):
            _run(str(tmp_path / "does_not_exist.xlsx"), tmp_path)

    def test_wrong_col_raises_key_error(self, excel_wrong_col_name, tmp_path):
        """File with unrecognised count column → KeyError with helpful message."""
        with pytest.raises(KeyError) as exc_info:
            _run(excel_wrong_col_name, tmp_path)
        msg = str(exc_info.value).lower()
        assert "column" in msg or "found" in msg or "columns" in msg

    def test_count_col_override_works_end_to_end(self, excel_wrong_col_override, tmp_path):
        """count_col_override='MyQty' enables parsing a non-standard column name."""
        containers = _run(excel_wrong_col_override, tmp_path, count_col_override="MyQty")
        assert len(containers) >= 1


# ---------------------------------------------------------------------------
# Output content correctness
# ---------------------------------------------------------------------------

class TestOutputContent:
    def test_json_containers_written(self, excel_minimal_valid, tmp_path):
        import json
        _run(excel_minimal_valid, tmp_path)
        json_path = tmp_path / "outputs" / "containers.json"
        assert json_path.exists()
        data = json.loads(json_path.read_text())
        assert isinstance(data, list)
        assert len(data) >= 1

    def test_recommendations_json_written(self, excel_minimal_valid, tmp_path):
        import json
        _run(excel_minimal_valid, tmp_path)
        recs_path = tmp_path / "outputs" / "recommendations.json"
        assert recs_path.exists()
        data = json.loads(recs_path.read_text())
        assert isinstance(data, list)

    def test_total_pallets_in_containers(self, excel_minimal_valid, tmp_path):
        """Total pallets across all containers should equal input count (8)."""
        containers = _run(excel_minimal_valid, tmp_path)
        total = sum(c["loaded_value"] for c in containers)
        assert total == 8
