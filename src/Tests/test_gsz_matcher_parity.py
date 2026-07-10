#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Регресс-тест: якорный предфильтр vs полный перебор."""

from __future__ import annotations

import json
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from gsz_matcher_parallel import (  # noqa: E402
    DEFAULT_AND_FULL,
    DEFAULT_AND_NOT,
    DEFAULT_OR_FULL,
    DEFAULT_OR_NOT,
    build_meta_row,
    match_single_holding,
    match_single_holding_brute,
    read_excel_table,
    worker_init,
)


class TestGszMatcherParity(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        config_path = ROOT / "config.json"
        with config_path.open(encoding="utf-8") as f:
            cfg = json.load(f)
        block = cfg["gsz_matcher_parallel"]
        cls.input_xlsx = ROOT / block["input_xlsx"]
        cls.holding_table = block["holding_table"]
        cls.base_table = block["base_table"]
        cls.holding_column = block["holding_column"]
        cls.gsz_column = block["gsz_column"]
        cls.and_full_cols = block.get("and_full_cols", DEFAULT_AND_FULL)
        cls.and_not_cols = block.get("and_not_cols", DEFAULT_AND_NOT)
        cls.or_full_cols = block.get("or_full_cols", DEFAULT_OR_FULL)
        cls.or_not_cols = block.get("or_not_cols", DEFAULT_OR_NOT)

        base_rows = read_excel_table(cls.input_xlsx, cls.base_table)
        metas = tuple(
            build_meta_row(
                row=r,
                gsz_col=cls.gsz_column,
                and_full_cols=cls.and_full_cols,
                and_not_cols=cls.and_not_cols,
                or_full_cols=cls.or_full_cols,
                or_not_cols=cls.or_not_cols,
            )
            for r in base_rows
        )
        worker_init(metas)

        hold_rows = read_excel_table(cls.input_xlsx, cls.holding_table)
        cls.holding_texts = [r.get(cls.holding_column) for r in hold_rows]

    def test_fast_matches_brute_on_all_holdings(self) -> None:
        mismatches: list[tuple[int, str, tuple, tuple]] = []
        for i, text in enumerate(self.holding_texts):
            fast = match_single_holding(text)
            brute = match_single_holding_brute(text)
            if fast != brute:
                mismatches.append((i, str(text), fast, brute))
                if len(mismatches) >= 5:
                    break

        if mismatches:
            sample = mismatches[0]
            self.fail(
                f"Расхождений fast vs brute: {len(mismatches)}+; "
                f"пример idx={sample[0]} text={sample[1]!r} fast={sample[2]} brute={sample[3]}"
            )


if __name__ == "__main__":
    unittest.main()
