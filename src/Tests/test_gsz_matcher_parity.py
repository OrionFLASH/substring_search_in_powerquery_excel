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
    DEFAULT_BASE_OUTPUT_COLUMNS,
    DEFAULT_HOLDING_OUTPUT_COLUMNS,
    DEFAULT_MATCH_STATUS_TEXTS,
    build_base_holding_match_columns,
    build_meta_row,
    enrich_base_rows,
    match_single_holding,
    match_single_holding_brute,
    normalize_holding_id,
    output_column_names,
    parse_sheet_output_columns,
    read_excel_table,
    reorder_base_row_columns,
    resolve_output_format,
    row_matches,
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
        if not cls.input_xlsx.exists():
            raise unittest.SkipTest(f"Файл тестовой книги не найден: {cls.input_xlsx}")
        cls.holding_table = block["holding_table"]
        cls.base_table = block["base_table"]
        cls.holding_column = block["holding_column"]
        cls.gsz_column = block["gsz_column"]
        from gsz_matcher_parallel import (  # noqa: E402
            DEFAULT_AND_FULL,
            DEFAULT_AND_NOT,
            DEFAULT_AND_NON,
            DEFAULT_OR_FULL,
            DEFAULT_OR_NOT,
            DEFAULT_OR_NON,
        )

        cls.and_full_cols = block.get("and_full_cols", DEFAULT_AND_FULL)
        cls.and_not_cols = block.get("and_not_cols", DEFAULT_AND_NOT)
        cls.and_non_cols = block.get("and_non_cols", DEFAULT_AND_NON)
        cls.or_full_cols = block.get("or_full_cols", DEFAULT_OR_FULL)
        cls.or_not_cols = block.get("or_not_cols", DEFAULT_OR_NOT)
        cls.or_non_cols = block.get("or_non_cols", DEFAULT_OR_NON)

        base_rows = read_excel_table(cls.input_xlsx, cls.base_table)
        hold_rows = read_excel_table(cls.input_xlsx, cls.holding_table)
        cls.holding_ids = [
            normalize_holding_id(r.get(block.get("holding_id_column", "ID холдинга")))
            for r in hold_rows
        ]
        holdings_id_set = frozenset(h for h in cls.holding_ids if h)
        fix_id_col = block.get("fix_id_col", "key_fix_id")
        metas = tuple(
            build_meta_row(
                row=r,
                gsz_col=cls.gsz_column,
                and_full_cols=cls.and_full_cols,
                and_not_cols=cls.and_not_cols,
                and_non_cols=cls.and_non_cols,
                or_full_cols=cls.or_full_cols,
                or_not_cols=cls.or_not_cols,
                or_non_cols=cls.or_non_cols,
                fix_id_col=fix_id_col,
                holdings_id_set=holdings_id_set,
            )
            for r in base_rows
        )
        worker_init(metas)

        cls.holding_texts = [r.get(cls.holding_column) for r in hold_rows]

    def test_fast_matches_brute_on_all_holdings(self) -> None:
        mismatches: list[tuple[int, str, tuple, tuple]] = []
        for i, text in enumerate(self.holding_texts):
            fast = match_single_holding(self.holding_ids[i], text)
            brute = match_single_holding_brute(self.holding_ids[i], text)
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


class TestBaseHoldingColumns(unittest.TestCase):
    def test_non_tokens_and_exclude_when_all_found_without_spaces(self) -> None:
        meta = build_meta_row(
            row={
                "Наименование, регион": "X",
                "key_and_full_1": "пик",
                "key_and_not_1": "",
                "key_and_non_1": "ммб",
                "key_and_non_2": "ну",
                "key_or_full_1": "",
                "key_or_not_1": "",
                "key_or_non_1": "",
            },
            gsz_col="Наименование, регион",
            and_full_cols=["key_and_full_1"],
            and_not_cols=["key_and_not_1"],
            and_non_cols=["key_and_non_1", "key_and_non_2"],
            or_full_cols=["key_or_full_1"],
            or_not_cols=["key_or_not_1"],
            or_non_cols=["key_or_non_1"],
        )
        self.assertFalse(row_matches("Ну ПИК ММБ", meta))
        self.assertTrue(row_matches("Ну ПИК", meta))

    def test_non_tokens_or_exclude_when_any_found_without_spaces(self) -> None:
        meta = build_meta_row(
            row={
                "Наименование, регион": "X",
                "key_and_full_1": "пик",
                "key_and_not_1": "",
                "key_and_non_1": "",
                "key_or_full_1": "",
                "key_or_not_1": "",
                "key_or_non_1": "ммб",
            },
            gsz_col="Наименование, регион",
            and_full_cols=["key_and_full_1"],
            and_not_cols=["key_and_not_1"],
            and_non_cols=["key_and_non_1"],
            or_full_cols=["key_or_full_1"],
            or_not_cols=["key_or_not_1"],
            or_non_cols=["key_or_non_1"],
        )
        self.assertFalse(row_matches("ну пикммб", meta))
        self.assertFalse(row_matches("ну пик ммб", meta))
        self.assertTrue(row_matches("ну пик", meta))

    def test_build_base_holding_match_columns_single(self) -> None:
        hold_rows = [
            {"ID холдинга": 101, "Холдинг": "ГК ПИК"},
            {"ID холдинга": 202, "Холдинг": "Самолет"},
        ]
        primary, debug = build_base_holding_match_columns(
            matched_pairs=[(0, False)],
            hold_rows=hold_rows,
            holding_id_column="ID холдинга",
            holding_name_column="Холдинг",
            texts=DEFAULT_MATCH_STATUS_TEXTS,
        )
        self.assertEqual(primary, "[101]: ГК ПИК")
        self.assertEqual(debug, "[101]: ГК ПИК;")

    def test_build_base_holding_match_columns_multiple(self) -> None:
        hold_rows = [
            {"ID холдинга": 101, "Холдинг": "ГК ПИК"},
            {"ID холдинга": 202, "Холдинг": "Самолет"},
        ]
        primary, debug = build_base_holding_match_columns(
            matched_pairs=[(0, False), (1, False)],
            hold_rows=hold_rows,
            holding_id_column="ID холдинга",
            holding_name_column="Холдинг",
            texts=DEFAULT_MATCH_STATUS_TEXTS,
        )
        self.assertEqual(primary, ":=>")
        self.assertEqual(debug, "[101]: ГК ПИК;\n[202]: Самолет;")

    def test_enrich_base_rows_column_order(self) -> None:
        base_rows = [{"Наименование, регион": "ПИК, Москва", "key_and_full_1": "пик"}]
        hold_rows = [{"ID холдинга": 7, "Холдинг": "ГК ПИК"}]
        base_columns = parse_sheet_output_columns(
            sheet_cfg=None,
            defaults=DEFAULT_BASE_OUTPUT_COLUMNS,
            default_width=30,
        )
        enrich_base_rows(
            base_rows=base_rows,
            base_metas=[
                build_meta_row(
                    row=base_rows[0],
                    gsz_col="Наименование, регион",
                    and_full_cols=["key_and_full_1"],
                    and_not_cols=[],
                    and_non_cols=[],
                    or_full_cols=[],
                    or_not_cols=[],
                    or_non_cols=[],
                )
            ],
            all_key_cols=["key_and_full_1"],
            per_row_holding_counts=[1],
            per_row_matched_holding_indices=[[(0, False)]],
            hold_rows=hold_rows,
            holding_id_column="ID холдинга",
            holding_name_column="Холдинг",
            base_columns=base_columns,
            status_texts=DEFAULT_MATCH_STATUS_TEXTS,
        )
        self.assertEqual(
            list(base_rows[0].keys())[-len(base_columns) :],
            output_column_names(base_columns),
        )
        self.assertEqual(base_rows[0]["найденный холдинг"], "[7]: ГК ПИК")
        self.assertEqual(base_rows[0]["Отладка_найденного_холдинга"], "[7]: ГК ПИК;")

    def test_reorder_base_row_columns(self) -> None:
        base_columns = DEFAULT_BASE_OUTPUT_COLUMNS
        row = {
            "Наименование, регион": "X",
            "число повторов": 1,
            "кол-во холдингов": 2,
            "найденный холдинг": "a",
            "Отладка_найденного_холдинга": "b",
            "статус": "найдено соответствие",
            "строка ключа": "k",
            "длина ключа": 1,
        }
        ordered = reorder_base_row_columns(row, base_columns)
        self.assertEqual(
            list(ordered.keys()),
            ["Наименование, регион", *output_column_names(base_columns)],
        )

    def test_resolve_output_format_custom_names(self) -> None:
        resolved = resolve_output_format(
            {
                "min_width_all": 40,
                "holding_sheet": {
                    "columns": {
                        "gsz_primary": {"name": "ГСЗ", "width": 120},
                        "gsz_debug": {"name": "Отладка ГСЗ", "width": 80, "wrap": False},
                        "match_status": {"name": "Статус", "width": 40},
                        "match_count": {"name": "Совпадений", "width": 25},
                    }
                },
                "base_sheet": {
                    "columns": {
                        "holding_count": {"name": "Холдингов", "width": 35},
                        "found_holding": {"name": "Холдинг", "width": 140},
                        "found_holding_debug": {"name": "Отладка холдинга", "width": 90, "wrap": True},
                        "match_status": {"name": "Статус базы", "width": 40},
                        "key_string": {"name": "Ключ", "width": 35},
                        "key_length": {"name": "Длина", "width": 35},
                        "key_repeat_count": {"name": "Повторы", "width": 35},
                    }
                },
            }
        )
        holding_by_key = {column.key: column for column in resolved["holding_columns"]}
        base_by_key = {column.key: column for column in resolved["base_columns"]}
        self.assertEqual(holding_by_key["gsz_primary"].name, "ГСЗ")
        self.assertEqual(holding_by_key["gsz_primary"].width, 120)
        self.assertEqual(base_by_key["found_holding_debug"].wrap, True)
        self.assertEqual(holding_by_key["match_status"].wrap, True)
        self.assertEqual(base_by_key["match_status"].wrap, True)
        self.assertEqual(resolved["min_width_all"], 40)

    def test_resolve_output_format_rejects_unknown_key(self) -> None:
        with self.assertRaises(ValueError):
            resolve_output_format(
                {
                    "holding_sheet": {
                        "columns": {
                            "unknown_key": {"name": "X", "width": 10},
                        }
                    }
                }
            )


class TestOrPrefilterAnchors(unittest.TestCase):
    """Регресс: OR-only с латиницей+кириллицей не должен теряться на предфильтре."""

    _EMPTY_AND = {
        "key_and_full_1": "",
        "key_and_full_2": "",
        "key_and_full_3": "",
        "key_and_not_1": "",
        "key_and_not_2": "",
        "key_and_not_3": "",
        "key_and_non_1": "",
        "key_and_non_2": "",
        "key_or_non_1": "",
        "key_or_non_2": "",
        "key_fix_id": "",
    }

    def _meta(self, **or_keys: str):
        row = {
            "Наименование, регион": "Seven Suns GSZ",
            **self._EMPTY_AND,
            "key_or_full_1": "",
            "key_or_full_2": "",
            "key_or_full_3": "",
            "key_or_not_1": "",
            "key_or_not_2": "",
            "key_or_not_3": "",
        }
        row.update(or_keys)
        return build_meta_row(
            row=row,
            gsz_col="Наименование, регион",
            and_full_cols=["key_and_full_1", "key_and_full_2", "key_and_full_3"],
            and_not_cols=["key_and_not_1", "key_and_not_2", "key_and_not_3"],
            and_non_cols=["key_and_non_1", "key_and_non_2"],
            or_full_cols=["key_or_full_1", "key_or_full_2", "key_or_full_3"],
            or_not_cols=["key_or_not_1", "key_or_not_2", "key_or_not_3"],
            or_non_cols=["key_or_non_1", "key_or_non_2"],
        )

    def test_or_only_latin_cyrillic_fast_equals_brute(self) -> None:
        meta = self._meta(
            key_or_full_1="Suns",
            key_or_full_2="САНС",
            key_or_not_1="Seven",
            key_or_not_2="СЕВЕН",
        )
        holding = "СЕВЕН САНС ДЕВЕЛОПМЕНТ"
        self.assertTrue(row_matches(holding, meta))
        # Все OR-токены в предфильтре (не один латинский якорь)
        words = {t.word for t in meta.anchors}
        self.assertEqual(words, {"suns", "санс", "seven", "севен"})

        worker_init((meta,), DEFAULT_MATCH_STATUS_TEXTS)
        fast = match_single_holding("1", holding)
        brute = match_single_holding_brute("1", holding)
        self.assertEqual(fast, brute)
        self.assertEqual(fast.count, 1)
        self.assertEqual(fast.primary, "Seven Suns GSZ")

    def test_and_still_uses_single_anchor(self) -> None:
        meta = build_meta_row(
            row={
                "Наименование, регион": "AND GSZ",
                "key_and_full_1": "санс",
                "key_and_not_1": "севен",
                "key_or_full_1": "suns",
                "key_or_not_1": "seven",
            },
            gsz_col="Наименование, регион",
            and_full_cols=["key_and_full_1"],
            and_not_cols=["key_and_not_1"],
            and_non_cols=[],
            or_full_cols=["key_or_full_1"],
            or_not_cols=["key_or_not_1"],
            or_non_cols=[],
        )
        # При AND предфильтр остаётся одним обязательным якорем (длиннейший not)
        self.assertEqual(len(meta.anchors), 1)
        self.assertEqual(meta.anchors[0].word, "севен")
        self.assertFalse(meta.anchors[0].is_full)


if __name__ == "__main__":
    unittest.main()
