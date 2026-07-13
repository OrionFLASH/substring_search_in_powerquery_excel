#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Тесты key_fix_id и колонки «статус»."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from gsz_matcher_parallel import (  # noqa: E402
    STATUS_FIX_NOT_FOUND,
    STATUS_FIXED,
    STATUS_MULTIPLE,
    STATUS_NONE,
    STATUS_SINGLE,
    BaseMeta,
    build_base_holding_match_columns,
    build_meta_row,
    compute_base_row_status,
    compute_holding_status,
    enrich_base_rows,
    format_match_columns,
    match_single_holding,
    match_single_holding_brute,
    parse_sheet_output_columns,
    parse_fix_ids,
    resolve_fix_mode,
    row_matches,
    worker_init,
    DEFAULT_BASE_OUTPUT_COLUMNS,
)


class TestKeyFixIdHelpers(unittest.TestCase):
    def test_parse_fix_ids_single_and_multiple(self) -> None:
        self.assertEqual(parse_fix_ids("100"), ("100",))
        self.assertEqual(parse_fix_ids("100; 200"), ("100", "200"))
        self.assertEqual(parse_fix_ids("100;200"), ("100", "200"))
        self.assertEqual(parse_fix_ids(None), ())

    def test_resolve_fix_mode(self) -> None:
        holdings = frozenset({"100", "200"})
        self.assertEqual(resolve_fix_mode((), holdings), "none")
        self.assertEqual(resolve_fix_mode(("100",), holdings), "resolved")
        self.assertEqual(resolve_fix_mode(("999",), holdings), "fallback")

    def test_compute_holding_status(self) -> None:
        self.assertEqual(compute_holding_status(0, 0), STATUS_NONE)
        self.assertEqual(compute_holding_status(1, 0), STATUS_SINGLE)
        self.assertEqual(compute_holding_status(2, 0), STATUS_MULTIPLE)
        self.assertEqual(compute_holding_status(1, 1), STATUS_FIXED)
        self.assertEqual(compute_holding_status(2, 2), STATUS_FIXED)

    def test_format_match_columns_no_status_phrase(self) -> None:
        primary, debug = format_match_columns(["A", "B"])
        self.assertEqual(primary, "A")
        self.assertEqual(debug, "A;\nB")


class TestKeyFixIdScenarios(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.holdings_id_set = frozenset({"100", "200"})

    def _meta(
        self,
        gsz: str,
        and_full: str = "",
        fix_id: str = "",
    ) -> BaseMeta:
        row = {
            "Наименование, регион": gsz,
            "key_and_full_1": and_full,
            "key_and_not_1": "",
            "key_and_non_1": "",
            "key_or_full_1": "",
            "key_or_not_1": "",
            "key_or_non_1": "",
            "key_fix_id": fix_id,
        }
        return build_meta_row(
            row=row,
            gsz_col="Наименование, регион",
            and_full_cols=["key_and_full_1"],
            and_not_cols=["key_and_not_1"],
            and_non_cols=["key_and_non_1"],
            or_full_cols=["key_or_full_1"],
            or_not_cols=["key_or_not_1"],
            or_non_cols=["key_or_non_1"],
            holdings_id_set=self.holdings_id_set,
        )

    def test_f1_fixed_match(self) -> None:
        metas = (self._meta("ГСЗ Fix", fix_id="100"),)
        worker_init(metas)
        result = match_single_holding("100", "любой текст")
        self.assertEqual(result.status, STATUS_FIXED)
        self.assertEqual(result.primary, "ГСЗ Fix")
        self.assertFalse(row_matches("другой текст", metas[0]))

    def test_f2_fix_missing_fallback_single_key(self) -> None:
        metas = (self._meta("ГСЗ Key", and_full="пик", fix_id="999"),)
        worker_init(metas)
        result = match_single_holding("100", "ГК ПИК")
        self.assertEqual(result.status, STATUS_SINGLE)
        self.assertEqual(result.primary, "ГСЗ Key")

    def test_f3_fix_missing_no_key_match(self) -> None:
        metas = (self._meta("ГСЗ Key", and_full="самолет", fix_id="999"),)
        worker_init(metas)
        base_meta = metas[0]
        self.assertEqual(compute_base_row_status(base_meta, []), STATUS_FIX_NOT_FOUND)

    def test_f4_multiple_key_matches_status_only(self) -> None:
        metas = (
            self._meta("ГСЗ 1", and_full="пик"),
            self._meta("ГСЗ 2", and_full="пик"),
        )
        worker_init(metas)
        result = match_single_holding("100", "ГК ПИК")
        self.assertEqual(result.status, STATUS_MULTIPLE)
        self.assertEqual(result.primary, "ГСЗ 1")
        self.assertIn("ГСЗ 2", result.debug)

    def test_fast_equals_brute_with_fix(self) -> None:
        metas = (
            self._meta("Fix GSZ", fix_id="100"),
            self._meta("Key GSZ", and_full="пик", fix_id="999"),
        )
        worker_init(metas)
        for holding_id, text in [("100", "x"), ("200", "ГК ПИК"), ("300", "нет")]:
            fast = match_single_holding(holding_id, text)
            brute = match_single_holding_brute(holding_id, text)
            self.assertEqual(fast, brute, f"holding_id={holding_id} text={text!r}")


class TestKeyFixIdBaseProjection(unittest.TestCase):
    def test_build_base_multiple_primary_is_first_entry(self) -> None:
        hold_rows = [
            {"ID холдинга": 101, "Холдинг": "ГК ПИК"},
            {"ID холдинга": 202, "Холдинг": "Самолет"},
        ]
        primary, debug = build_base_holding_match_columns(
            matched_holding_indices=[0, 1],
            hold_rows=hold_rows,
            holding_id_column="ID холдинга",
            holding_name_column="Холдинг",
        )
        self.assertEqual(primary, "[101]: ГК ПИК")
        self.assertIn("[202]: Самолет", debug)

    def test_enrich_base_rows_writes_status(self) -> None:
        base_rows = [{"Наименование, регион": "ПИК", "key_and_full_1": "пик", "key_fix_id": "999"}]
        hold_rows = [{"ID холдинга": 7, "Холдинг": "ГК ПИК"}]
        base_columns = parse_sheet_output_columns(
            sheet_cfg=None,
            defaults=DEFAULT_BASE_OUTPUT_COLUMNS,
            default_width=30,
        )
        meta = build_meta_row(
            row=base_rows[0],
            gsz_col="Наименование, регион",
            and_full_cols=["key_and_full_1"],
            and_not_cols=[],
            and_non_cols=[],
            or_full_cols=[],
            or_not_cols=[],
            or_non_cols=[],
            holdings_id_set=frozenset({"7"}),
        )
        enrich_base_rows(
            base_rows=base_rows,
            base_metas=[meta],
            all_key_cols=["key_and_full_1"],
            per_row_holding_counts=[1],
            per_row_matched_holding_indices=[[(0, False)]],
            hold_rows=hold_rows,
            holding_id_column="ID холдинга",
            holding_name_column="Холдинг",
            base_columns=base_columns,
        )
        self.assertEqual(base_rows[0]["найденный холдинг"], "[7]: ГК ПИК")
        self.assertEqual(base_rows[0]["статус"], STATUS_SINGLE)


if __name__ == "__main__":
    unittest.main()
