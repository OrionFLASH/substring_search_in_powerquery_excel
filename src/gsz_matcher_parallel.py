#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Параллельное сопоставление холдингов с условными ГСЗ.

Скрипт переносит логику Power Query в Python:
- and/or блоки;
- full/not режимы;
- непересечение интервалов для AND.

Оптимизации:
- предобработка справочника _base_gsz;
- якорный предфильтр кандидатов;
- параллельная обработка строк _HOLD_OD.
"""

from __future__ import annotations

import argparse
import json
import multiprocessing as mp
import re
import time
from concurrent.futures import ProcessPoolExecutor
from concurrent.futures import wait
from concurrent.futures import FIRST_COMPLETED
from dataclasses import dataclass
from pathlib import Path
from typing import Any


DEFAULT_AND_FULL = ["key_and_full_1", "key_and_full_2", "key_and_full_3"]
DEFAULT_AND_NOT = ["key_and_not_1", "key_and_not_2", "key_and_not_3"]
DEFAULT_OR_FULL = ["key_or_full_1", "key_or_full_2", "key_or_full_3"]
DEFAULT_OR_NOT = ["key_or_not_1", "key_or_not_2", "key_or_not_3"]
DEFAULT_CONFIG_PATH = Path(__file__).resolve().parent.parent / "config.json"


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def is_letter(ch: str) -> bool:
    if len(ch) != 1:
        return False
    c = ch.lower()
    return ("a" <= c <= "z") or ("а" <= c <= "я") or c == "ё"


def all_positions_full(text: str, word: str) -> list[tuple[int, int]]:
    if not word or len(word) > len(text):
        return []
    out: list[tuple[int, int]] = []
    start = 0
    wl = len(word)
    while True:
        pos = text.find(word, start)
        if pos == -1:
            break
        out.append((pos, pos + wl - 1))
        start = pos + 1
    return out


def positions_not(text: str, word: str) -> list[tuple[int, int]]:
    result: list[tuple[int, int]] = []
    n = len(text)
    for s, e in all_positions_full(text, word):
        left_ok = s == 0 or not is_letter(text[s - 1])
        right_ok = e == n - 1 or not is_letter(text[e + 1])
        if left_ok and right_ok:
            result.append((s, e))
    return result


def overlap(a: tuple[int, int], b: tuple[int, int]) -> bool:
    return not (a[1] < b[0] or b[1] < a[0])


def and_non_overlapping(position_lists: list[list[tuple[int, int]]], idx: int = 0, chosen: list[tuple[int, int]] | None = None) -> bool:
    chosen = chosen or []
    if idx >= len(position_lists):
        return True
    for interval in position_lists[idx]:
        if any(overlap(interval, prev) for prev in chosen):
            continue
        if and_non_overlapping(position_lists, idx + 1, chosen + [interval]):
            return True
    return False


def extract_words(text: str) -> set[str]:
    # Для fast-предфильтра по режиму "not".
    return set(re.findall(r"[A-Za-zА-Яа-яЁё0-9]+", text.lower()))


@dataclass(frozen=True)
class Token:
    word: str
    is_full: bool


@dataclass(frozen=True)
class BaseMeta:
    gsz_value: str
    has_keys: bool
    and_tokens: tuple[Token, ...]
    or_tokens: tuple[Token, ...]
    anchor: Token | None


def parse_cols(value: str) -> list[str]:
    if isinstance(value, list):
        return [str(v).strip() for v in value if str(v).strip()]
    return [v.strip() for v in str(value).split(",") if v.strip()]


def pick_best_anchor(tokens: list[Token]) -> Token | None:
    if not tokens:
        return None
    return max(tokens, key=lambda t: len(t.word))


def pick_anchor(and_tokens: list[Token], or_tokens: list[Token]) -> Token | None:
    and_not = [t for t in and_tokens if not t.is_full]
    and_full = [t for t in and_tokens if t.is_full]
    or_not = [t for t in or_tokens if not t.is_full]
    or_full = [t for t in or_tokens if t.is_full]
    return (
        pick_best_anchor(and_not)
        or pick_best_anchor(and_full)
        or pick_best_anchor(or_not)
        or pick_best_anchor(or_full)
    )


def build_meta_row(
    row: dict[str, Any],
    gsz_col: str,
    and_full_cols: list[str],
    and_not_cols: list[str],
    or_full_cols: list[str],
    or_not_cols: list[str],
) -> BaseMeta:
    and_full_raw = [normalize_text(row.get(c, "")) for c in and_full_cols]
    and_not_raw = [normalize_text(row.get(c, "")) for c in and_not_cols]
    or_full_raw = [normalize_text(row.get(c, "")) for c in or_full_cols]
    or_not_raw = [normalize_text(row.get(c, "")) for c in or_not_cols]

    and_tokens = [Token(w, True) for w in and_full_raw if w] + [Token(w, False) for w in and_not_raw if w]
    or_tokens = [Token(w, True) for w in or_full_raw if w] + [Token(w, False) for w in or_not_raw if w]
    has_keys = bool(and_tokens or or_tokens)
    anchor = pick_anchor(and_tokens, or_tokens)

    gsz_value = str(row.get(gsz_col, "") or "").strip()
    return BaseMeta(
        gsz_value=gsz_value,
        has_keys=has_keys,
        and_tokens=tuple(and_tokens),
        or_tokens=tuple(or_tokens),
        anchor=anchor,
    )


def row_matches(text: str, meta: BaseMeta) -> bool:
    if not text or not meta.has_keys:
        return False

    # AND
    if meta.and_tokens:
        position_lists: list[list[tuple[int, int]]] = []
        for t in meta.and_tokens:
            pos = all_positions_full(text, t.word) if t.is_full else positions_not(text, t.word)
            if not pos:
                return False
            position_lists.append(pos)
        if not and_non_overlapping(position_lists):
            return False

    # OR
    if meta.or_tokens:
        ok_or = False
        for t in meta.or_tokens:
            pos = all_positions_full(text, t.word) if t.is_full else positions_not(text, t.word)
            if pos:
                ok_or = True
                break
        if not ok_or:
            return False

    return True


def read_excel_table(path: Path, table_name: str) -> list[dict[str, Any]]:
    from openpyxl import load_workbook
    from openpyxl.utils.cell import range_boundaries

    wb = load_workbook(path, data_only=True, read_only=False)
    for ws in wb.worksheets:
        if table_name in ws.tables:
            table = ws.tables[table_name]
            min_col, min_row, max_col, max_row = range_boundaries(table.ref)
            rows: list[list[Any]] = []
            for row in ws.iter_rows(
                min_row=min_row,
                max_row=max_row,
                min_col=min_col,
                max_col=max_col,
                values_only=True,
            ):
                rows.append(list(row))
            if not rows:
                return []
            headers = [str(h) if h is not None else "" for h in rows[0]]
            body = rows[1:]
            out: list[dict[str, Any]] = []
            for row in body:
                rec = {headers[i]: row[i] if i < len(row) else None for i in range(len(headers))}
                out.append(rec)
            return out
    raise ValueError(f"Таблица '{table_name}' не найдена в {path}")


BASE_METAS: tuple[BaseMeta, ...] = ()


def worker_init(base_metas: tuple[BaseMeta, ...]) -> None:
    global BASE_METAS
    BASE_METAS = base_metas


def match_single_holding(text_value: Any) -> tuple[str, str]:
    text = normalize_text(text_value)
    if not text:
        return "-", "-"
    words = extract_words(text)

    matches: list[str] = []
    for meta in BASE_METAS:
        a = meta.anchor
        if a is not None:
            if a.is_full:
                if a.word not in text:
                    continue
            else:
                if a.word not in words:
                    continue
        if row_matches(text, meta):
            if meta.gsz_value:
                matches.append(meta.gsz_value)

    if not matches:
        return "-", "-"
    return matches[0], "; ".join(matches)


def ensure_columns(rows: list[dict[str, Any]], cols: list[str], where: str) -> None:
    if not rows:
        raise ValueError(f"{where} пуста")
    available = set(rows[0].keys())
    missing = [c for c in cols if c not in available]
    if missing:
        raise ValueError(f"В {where} отсутствуют колонки: {missing}")


def write_sheet(ws: Any, rows: list[dict[str, Any]]) -> None:
    if not rows:
        return
    headers = list(rows[0].keys())
    ws.append(headers)
    for row in rows:
        ws.append([row.get(h) for h in headers])


def write_output_xlsx(
    output_path: Path,
    holding_rows: list[dict[str, Any]],
    base_rows: list[dict[str, Any]],
    holding_sheet: str,
    base_sheet: str,
) -> None:
    from openpyxl import Workbook

    wb = Workbook()
    ws1 = wb.active
    ws1.title = holding_sheet[:31] if holding_sheet else "HOLD_OD"
    write_sheet(ws1, holding_rows)

    ws2 = wb.create_sheet(title=base_sheet[:31] if base_sheet else "base_gsz")
    write_sheet(ws2, base_rows)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def chunked(seq: list[Any], size: int) -> list[list[Any]]:
    if size <= 0:
        size = 1
    return [seq[i : i + size] for i in range(0, len(seq), size)]


def match_holding_batch(text_batch: list[Any]) -> list[tuple[str, str]]:
    return [match_single_holding(x) for x in text_batch]


def short_text(value: Any, max_len: int = 80) -> str:
    s = str(value) if value is not None else ""
    s = s.replace("\n", " ").replace("\r", " ").strip()
    if len(s) <= max_len:
        return s
    return s[: max_len - 3] + "..."


def log(message: str) -> None:
    print(message, flush=True)


def build_progress_message(
    processed: int,
    total_holdings: int,
    batch_idx: int,
    total_batches: int,
    match_started: float,
    holding_texts: list[Any],
    show_current_holding: bool,
    prefix: str = "[progress]",
) -> str:
    elapsed = time.perf_counter() - match_started
    speed = processed / elapsed if elapsed > 0 else 0.0
    remaining = max(0, total_holdings - processed)
    eta_sec = (remaining / speed) if speed > 0 else 0.0
    pct = (processed / total_holdings * 100.0) if total_holdings else 100.0
    msg = (
        f"{prefix} {processed}/{total_holdings} ({pct:.1f}%), "
        f"batch={batch_idx}/{total_batches}, speed={speed:.1f} hold/s, "
        f"eta={eta_sec:.1f}s"
    )
    if show_current_holding and processed > 0:
        msg += f", current='{short_text(holding_texts[processed - 1])}'"
    return msg


def build_batch_message(
    completed_batches: int,
    total_batches: int,
    work_batch_size: int,
    total_holdings: int,
    pending_futures: int,
) -> str:
    approx_done = min(total_holdings, completed_batches * work_batch_size)
    pct = (approx_done / total_holdings * 100.0) if total_holdings else 100.0
    return (
        f"[progress-batch] batches={completed_batches}/{total_batches}, "
        f"approx_holdings_done~={approx_done}/{total_holdings} ({pct:.1f}%), "
        f"in_flight={pending_futures}"
    )


def make_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Параллельное сопоставление холдингов и условных ГСЗ")
    p.add_argument("--input-xlsx", help="Путь к исходной Excel-книге")
    p.add_argument("--output-xlsx", help="Путь к выходному Excel-файлу")
    p.add_argument("--holding-table", help="Имя таблицы холдингов")
    p.add_argument("--base-table", help="Имя таблицы справочника ГСЗ")
    p.add_argument("--holding-column", help="Колонка с текстом холдинга")
    p.add_argument("--gsz-column", help="Колонка значения условного ГСЗ")
    p.add_argument("--and-full-cols", help="AND full колонки через запятую")
    p.add_argument("--and-not-cols", help="AND not колонки через запятую")
    p.add_argument("--or-full-cols", help="OR full колонки через запятую")
    p.add_argument("--or-not-cols", help="OR not колонки через запятую")
    p.add_argument("--workers", type=int, help="Число процессов")
    p.add_argument("--chunk-size", type=int, help="Размер чанка для process pool")
    p.add_argument("--config-json", default=str(DEFAULT_CONFIG_PATH), help="JSON-файл с параметрами")
    return p


def load_config(path: Path) -> dict[str, Any]:
    if not path.exists():
        raise FileNotFoundError(f"config-json не найден: {path}")
    with path.open(encoding="utf-8") as f:
        return json.load(f)


def resolve_settings(args: argparse.Namespace) -> dict[str, Any]:
    cfg = load_config(Path(args.config_json).expanduser().resolve())
    block = cfg.get("gsz_matcher_parallel", {})

    defaults: dict[str, Any] = {
        "input_xlsx": block.get("input_xlsx"),
        "output_xlsx": block.get("output_xlsx"),
        "holding_table": block.get("holding_table", "_HOLD_OD"),
        "base_table": block.get("base_table", "_base_gsz"),
        "holding_column": block.get("holding_column", "Холдинг"),
        "gsz_column": block.get("gsz_column", "Наименование, регион"),
        "and_full_cols": block.get("and_full_cols", DEFAULT_AND_FULL),
        "and_not_cols": block.get("and_not_cols", DEFAULT_AND_NOT),
        "or_full_cols": block.get("or_full_cols", DEFAULT_OR_FULL),
        "or_not_cols": block.get("or_not_cols", DEFAULT_OR_NOT),
        "workers": block.get("workers", max(1, (mp.cpu_count() or 2) - 1)),
        "chunk_size": block.get("chunk_size", 200),
        "work_batch_size": block.get("work_batch_size", 50),
        "log_stages": block.get("log_stages", True),
        "progress_every_holdings": block.get("progress_every_holdings", 1000),
        "progress_every_base_rows": block.get("progress_every_base_rows", 1000),
        "progress_every_batches": block.get("progress_every_batches", 25),
        "heartbeat_seconds": block.get("heartbeat_seconds", 10),
        "show_current_holding": block.get("show_current_holding", True),
    }

    # Обратная совместимость: можно использовать старые плоские ключи.
    if defaults["input_xlsx"] is None and cfg.get("input_xlsx"):
        defaults["input_xlsx"] = cfg.get("input_xlsx")
    if defaults["output_xlsx"] is None and cfg.get("output_xlsx"):
        defaults["output_xlsx"] = cfg.get("output_xlsx")

    cli_overrides = {
        "input_xlsx": args.input_xlsx,
        "output_xlsx": args.output_xlsx,
        "holding_table": args.holding_table,
        "base_table": args.base_table,
        "holding_column": args.holding_column,
        "gsz_column": args.gsz_column,
        "and_full_cols": args.and_full_cols,
        "and_not_cols": args.and_not_cols,
        "or_full_cols": args.or_full_cols,
        "or_not_cols": args.or_not_cols,
        "workers": args.workers,
        "chunk_size": args.chunk_size,
    }
    for key, value in cli_overrides.items():
        if value is not None:
            defaults[key] = value

    if not defaults["input_xlsx"] or not defaults["output_xlsx"]:
        raise ValueError(
            "Не заданы input/output пути для Python-матчера. "
            "Укажите их в config.json (блок gsz_matcher_parallel) "
            "или передайте --input-xlsx и --output-xlsx."
        )
    return defaults


def main() -> None:
    parser = make_arg_parser()
    args = parser.parse_args()
    settings = resolve_settings(args)

    input_xlsx = Path(settings["input_xlsx"]).expanduser().resolve()
    output_xlsx = Path(settings["output_xlsx"]).expanduser().resolve()

    and_full_cols = parse_cols(settings["and_full_cols"])
    and_not_cols = parse_cols(settings["and_not_cols"])
    or_full_cols = parse_cols(settings["or_full_cols"])
    or_not_cols = parse_cols(settings["or_not_cols"])

    t0 = time.perf_counter()
    if settings["log_stages"]:
        log("[stage] Запуск Python-матчера.")
        log(
            f"[stage] Конфиг: workers={settings['workers']}, chunk_size={settings['chunk_size']}, "
            f"work_batch_size={settings['work_batch_size']}, "
            f"progress_every={settings['progress_every_holdings']}, "
            f"heartbeat={settings['heartbeat_seconds']}s"
        )
    if settings["log_stages"]:
        log(f"[stage] Чтение таблицы {settings['holding_table']}...")
    hold_rows = read_excel_table(input_xlsx, settings["holding_table"])
    if settings["log_stages"]:
        log(f"[stage] Таблица {settings['holding_table']} загружена: {len(hold_rows)} строк.")
        log(f"[stage] Чтение таблицы {settings['base_table']}...")
    base_rows = read_excel_table(input_xlsx, settings["base_table"])
    if settings["log_stages"]:
        log(f"[stage] Таблица {settings['base_table']} загружена: {len(base_rows)} строк.")

    if settings["log_stages"]:
        log("[stage] Проверка обязательных колонок...")
    ensure_columns(hold_rows, [settings["holding_column"]], f"таблице {settings['holding_table']}")
    ensure_columns(
        base_rows,
        [settings["gsz_column"]] + and_full_cols + and_not_cols + or_full_cols + or_not_cols,
        f"таблице {settings['base_table']}",
    )

    if settings["log_stages"]:
        log("[stage] Подготовка метаданных справочника...")
    metas_list: list[BaseMeta] = []
    base_progress_every = max(1, int(settings["progress_every_base_rows"]))
    for idx, r in enumerate(base_rows, start=1):
        metas_list.append(
            build_meta_row(
                row=r,
                gsz_col=settings["gsz_column"],
                and_full_cols=and_full_cols,
                and_not_cols=and_not_cols,
                or_full_cols=or_full_cols,
                or_not_cols=or_not_cols,
            )
        )
        if settings["log_stages"] and (idx % base_progress_every == 0 or idx == len(base_rows)):
            log(f"[progress-base] {idx}/{len(base_rows)}")
    metas = tuple(metas_list)

    total_holdings = len(hold_rows)
    total_base = len(base_rows)
    approx_comparisons = total_holdings * total_base
    if settings["log_stages"]:
        log("[stage] Подготовка справочника завершена.")
        log(
            f"[stage] Оценка масштаба: {total_holdings} холдингов x "
            f"{total_base} строк справочника ~= {approx_comparisons} проверок"
        )

    holding_texts = [r.get(settings["holding_column"]) for r in hold_rows]
    pool_chunk_size = max(1, int(settings["chunk_size"]))
    work_batch_size = max(1, int(settings["work_batch_size"]))
    progress_every = max(1, int(settings["progress_every_holdings"]))
    progress_every_batches = max(1, int(settings["progress_every_batches"]))
    workers = max(1, int(settings["workers"]))
    heartbeat_seconds = max(1, int(settings["heartbeat_seconds"]))
    batches = chunked(holding_texts, work_batch_size)

    if settings["log_stages"]:
        log(
            f"[stage] Старт сопоставления: workers={workers}, work_batch={work_batch_size}, "
            f"pool_chunk={pool_chunk_size}, progress_every={progress_every}, "
            f"progress_every_batches={progress_every_batches}"
        )

    results: list[tuple[str, str]] = []
    processed = 0
    next_progress = progress_every
    match_started = time.perf_counter()
    with ProcessPoolExecutor(
        max_workers=workers,
        initializer=worker_init,
        initargs=(metas,),
    ) as ex:
        future_to_idx = {ex.submit(match_holding_batch, b): i for i, b in enumerate(batches)}
        ordered_batches: dict[int, list[tuple[str, str]]] = {}
        next_idx_to_flush = 0
        completed_batches = 0
        next_batch_progress = progress_every_batches
        pending = set(future_to_idx.keys())
        while pending:
            done, pending = wait(pending, timeout=heartbeat_seconds, return_when=FIRST_COMPLETED)
            if not done:
                log(
                    build_progress_message(
                        processed=processed,
                        total_holdings=total_holdings,
                        batch_idx=next_idx_to_flush,
                        total_batches=len(batches),
                        match_started=match_started,
                        holding_texts=holding_texts,
                        show_current_holding=settings["show_current_holding"],
                        prefix="[heartbeat]",
                    )
                )
                continue

            for fut in done:
                idx = future_to_idx[fut]
                ordered_batches[idx] = fut.result()
                completed_batches += 1
                if completed_batches >= next_batch_progress or completed_batches == len(batches):
                    log(
                        build_batch_message(
                            completed_batches=completed_batches,
                            total_batches=len(batches),
                            work_batch_size=work_batch_size,
                            total_holdings=total_holdings,
                            pending_futures=len(pending),
                        )
                    )
                    while next_batch_progress <= completed_batches:
                        next_batch_progress += progress_every_batches

            while next_idx_to_flush in ordered_batches:
                batch_result = ordered_batches.pop(next_idx_to_flush)
                results.extend(batch_result)
                processed += len(batch_result)

                if processed >= next_progress or processed == total_holdings:
                    log(
                        build_progress_message(
                            processed=processed,
                            total_holdings=total_holdings,
                            batch_idx=next_idx_to_flush + 1,
                            total_batches=len(batches),
                            match_started=match_started,
                            holding_texts=holding_texts,
                            show_current_holding=settings["show_current_holding"],
                            prefix="[progress]",
                        )
                    )
                    while next_progress <= processed:
                        next_progress += progress_every
                next_idx_to_flush += 1

    for row, res in zip(hold_rows, results):
        row["условное ГСЗ"] = res[0]
        row["Отладка_совпадения_ГСЗ"] = res[1]

    if settings["log_stages"]:
        log("[stage] Запись результата в Excel...")
    write_output_xlsx(
        output_path=output_xlsx,
        holding_rows=hold_rows,
        base_rows=base_rows,
        holding_sheet=settings["holding_table"],
        base_sheet=settings["base_table"],
    )

    t1 = time.perf_counter()
    if settings["log_stages"]:
        log("[stage] Готово.")
    log(f"Готово. Вход: {input_xlsx}")
    log(f"Результат: {output_xlsx}")
    log(f"Холдингов: {len(hold_rows)}, строк _base_gsz: {len(base_rows)}")
    log(
        f"Потоков: {max(1, int(settings['workers']))}, "
        f"work_batch: {max(1, int(settings['work_batch_size']))}, "
        f"pool_chunk: {max(1, int(settings['chunk_size']))}"
    )
    log(f"Время: {t1 - t0:.2f} сек")


if __name__ == "__main__":
    main()
