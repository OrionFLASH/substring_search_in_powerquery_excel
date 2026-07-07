#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Генерация ключей для таблицы _base_gsz из списка «Наименование, регион»."""

from __future__ import annotations

import json
import logging
import re
from dataclasses import dataclass, field
from itertools import combinations
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

# ─── Пути ───────────────────────────────────────────────────────────────────
КОРЕНЬ = Path(__file__).resolve().parent.parent
КОНФИГ_ПУТЬ = КОРЕНЬ / "config.json"
ВХОДНОЙ_ФАЙЛ = КОРЕНЬ / "КейЛОАД.txt"
ВЫХОДНОЙ_XLSX = КОРЕНЬ / "output" / "_base_gsz_keys.xlsx"

КОЛОНКИ_КЛЮЧЕЙ = [
    "key_and_full_1",
    "key_and_full_2",
    "key_and_full_3",
    "key_or_full_1",
    "key_or_full_2",
    "key_or_full_3",
    "key_and_not_1",
    "key_and_not_2",
    "key_and_not_3",
    "key_or_not_1",
    "key_or_not_2",
    "key_or_not_3",
]

# Служебные слова — не используются как ключи
СТОП_СЛОВА = {
    "гк",
    "группа",
    "группы",
    "холдинг",
    "ооо",
    "ао",
    "зао",
    "пао",
    "ск",
    "сз",
    "кп",
    "ак",
    "фонд",
    "компания",
    "корпорация",
    "концерн",
    "объединение",
    "застройщиков",
    "застройщик",
    "девелопмент",
    "development",
    "group",
    "holding",
    "строй",
    "строительный",
    "трест",
    "нпф",
    "мк",
    "нп",
    "ип",
}

# Слишком общие токены (встречаются у многих застройщиков)
ОБЩИЕ_ТОКЕНЫ = {
    "group",
    "development",
    "строй",
    "инвест",
    "дом",
    "сити",
    "плюс",
    "регион",
    "север",
    "юг",
    "восток",
    "запад",
    "центр",
    "новый",
    "новая",
    "строй",
    "проект",
}


@dataclass
class КлючиСтроки:
    """Набор ключей для одной строки справочника."""

    and_full: list[str] = field(default_factory=list)
    and_not: list[str] = field(default_factory=list)
    or_full: list[str] = field(default_factory=list)
    or_not: list[str] = field(default_factory=list)
    дубликат: bool = False
    комментарий: str = ""

    def в_словарь(self) -> dict[str, str]:
        """Преобразование в плоский словарь колонок Excel."""
        результат: dict[str, str] = {к: "" for к in КОЛОНКИ_КЛЮЧЕЙ}

        for i, слово in enumerate(self.and_full[:3], start=1):
            результат[f"key_and_full_{i}"] = слово
        for i, слово in enumerate(self.or_full[:3], start=1):
            результат[f"key_or_full_{i}"] = слово
        for i, слово in enumerate(self.and_not[:3], start=1):
            результат[f"key_and_not_{i}"] = слово
        for i, слово in enumerate(self.or_not[:3], start=1):
            результат[f"key_or_not_{i}"] = слово

        return результат


@dataclass
class ЗаписьГСЗ:
    """Одна строка справочника."""

    полное_имя: str
    короткое_имя: str
    токены: list[str]
    ключи: КлючиСтроки = field(default_factory=КлючиСтроки)


def настроить_логирование() -> None:
    """Настройка логгера INFO."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - [%(levelname)s] - %(message)s",
    )


def загрузить_конфиг() -> dict[str, Any]:
    """Загрузка config.json с значениями по умолчанию."""
    значения_по_умолчанию: dict[str, Any] = {
        "input_file": str(ВХОДНОЙ_ФАЙЛ),
        "output_file": str(ВЫХОДНОЙ_XLSX),
        "min_token_length": 2,
        "short_token_not_max_len": 5,
        "common_token_threshold": 80,
    }
    if КОНФИГ_ПУТЬ.exists():
        with КОНФИГ_ПУТЬ.open(encoding="utf-8") as файл:
            данные = json.load(файл)
        значения_по_умолчанию.update(данные)
    return значения_по_умолчанию


def нормализовать(текст: str) -> str:
    """Trim + нижний регистр."""
    return текст.strip().lower()


def это_буква(символ: str) -> bool:
    """Кириллица или латиница."""
    if len(символ) != 1:
        return False
    с = символ.lower()
    return ("a" <= с <= "z") or ("а" <= с <= "я") or с == "ё"


def все_позиции_full(текст: str, слово: str) -> list[tuple[int, int]]:
    """Интервалы вхождений подстроки (включительно по индексу конца)."""
    if not слово or not текст or len(слово) > len(текст):
        return []
    дл = len(слово)
    return [(i, i + дл - 1) for i in range(len(текст) - дл + 1) if текст[i : i + дл] == слово]


def позиции_not(текст: str, слово: str) -> list[tuple[int, int]]:
    """Интервалы отдельного слова (режим not)."""
    результат: list[tuple[int, int]] = []
    for старт, конец in все_позиции_full(текст, слово):
        слева = старт == 0 or not это_буква(текст[старт - 1])
        справа = конец == len(текст) - 1 or not это_буква(текст[конец + 1])
        if слева and справа:
            результат.append((старт, конец))
    return результат


def интервалы_пересекаются(а: tuple[int, int], б: tuple[int, int]) -> bool:
    """Проверка пересечения двух интервалов."""
    return not (а[1] < б[0] or б[1] < а[0])


def подобрать_без_пересечения(списки: list[list[tuple[int, int]]], индекс: int = 0, выбранные: list[tuple[int, int]] | None = None) -> bool:
    """Рекурсивный подбор позиций AND-токенов."""
    выбранные = выбранные or []
    if индекс >= len(списки):
        return True
    for интервал in списки[индекс]:
        if any(интервалы_пересекаются(интервал, в) for в in выбранные):
            continue
        if подобрать_без_пересечения(списки, индекс + 1, выбранные + [интервал]):
            return True
    return False


def проверить_and(текст: str, токены: list[tuple[str, bool]]) -> bool:
    """AND-блок: все токены найдены без пересечения."""
    if not токены:
        return True
    норм = нормализовать(текст)
    списки: list[list[tuple[int, int]]] = []
    for слово, is_full in токены:
        w = нормализовать(слово)
        if not w:
            return False
        позиции = все_позиции_full(норм, w) if is_full else позиции_not(норм, w)
        if not позиции:
            return False
        списки.append(позиции)
    return подобрать_без_пересечения(списки)


def проверить_or(текст: str, токены: list[tuple[str, bool]]) -> bool:
    """OR-блок: хотя бы один токен."""
    if not токены:
        return True
    норм = нормализовать(текст)
    for слово, is_full in токены:
        w = нормализовать(слово)
        if not w:
            continue
        позиции = все_позиции_full(норм, w) if is_full else позиции_not(норм, w)
        if позиции:
            return True
    return False


def строка_совпала(текст_холдинга: str, ключи: КлючиСтроки) -> bool:
    """Логика совпадения как в Power Query."""
    if not нормализовать(текст_холдинга):
        return False
    and_токены = [(w, True) for w in ключи.and_full] + [(w, False) for w in ключи.and_not]
    or_токены = [(w, True) for w in ключи.or_full] + [(w, False) for w in ключи.or_not]
    if not and_токены and not or_токены:
        return False
    return проверить_and(текст_холдинга, and_токены) and проверить_or(текст_холдинга, or_токены)


def извлечь_токены(короткое_имя: str, мин_длина: int, резервный: bool = False) -> list[str]:
    """Разбор имени на ключевые части (слова и части через дефис)."""
    текст = короткое_имя.strip()
    скобки = re.findall(r"\(([^)]+)\)", текст)
    текст = re.sub(r"\([^)]*\)", " ", текст)
    текст = re.sub(r"[®™]", "", текст)

    части = re.split(r"[\s\-–—/]+", текст)
    части.extend(re.split(r"[\s\-–—/]+", " ".join(скобки)))

    токены: list[str] = []
    стоп = СТОП_СЛОВА if not резервный else {"гк", "группа", "холдинг", "ооо", "ао", "зао", "пао"}
    общие = ОБЩИЕ_ТОКЕНЫ if not резервный else set()

    for часть in части:
        очищенная = re.sub(r"[^\w]", "", часть, flags=re.UNICODE)
        if len(очищенная) < мин_длина:
            continue
        if очищенная.lower() in стоп:
            continue
        if очищенная.lower() in общие:
            continue
        токены.append(очищенная)

    уникальные: list[str] = []
    видели: set[str] = set()
    for т in токены:
        кл = т.lower()
        if кл not in видели:
            видели.add(кл)
            уникальные.append(т)
    return уникальные


def частота_токенов(записи: list[ЗаписьГСЗ]) -> dict[str, int]:
    """Сколько строк содержат токен (по вхождению в короткое имя)."""
    частоты: dict[str, int] = {}
    for запись in записи:
        норм = нормализовать(запись.короткое_имя)
        for т in запись.токены:
            кл = т.lower()
            if кл in норм:
                частоты[кл] = частоты.get(кл, 0) + 1
    return частоты


def отсортировать_токены(токены: list[str], частоты: dict[str, int]) -> list[str]:
    """Редкие и длинные токены — в приоритете."""
    return sorted(
        токены,
        key=lambda т: (частоты.get(т.lower(), 9999), -len(т), т.lower()),
    )


def выбрать_режим_токена(токен: str, порог_not: int) -> bool:
    """True = full, False = not."""
    if len(токен) <= порог_not and токен.isalpha():
        return False
    return True


def построить_ключи_из_токенов(токены: list[tuple[str, bool]]) -> КлючиСтроки:
    """Сборка структуры ключей из списка (слово, is_full)."""
    ключи = КлючиСтроки()
    for слово, is_full in токены[:3]:
        if is_full:
            ключи.and_full.append(слово)
        else:
            ключи.and_not.append(слово)
    return ключи


def уникально_для_записи(ключи: КлючиСтроки, своя: ЗаписьГСЗ, все: list[ЗаписьГСЗ]) -> bool:
    """Ключи однозначно определяют запись по короткому имени."""
    if not строка_совпала(своя.короткое_имя, ключи):
        return False
    for другая in все:
        if другая is своя:
            continue
        if строка_совпала(другая.короткое_имя, ключи):
            return False
    return True


def подобрать_ключи(
    запись: ЗаписьГСЗ,
    все: list[ЗаписьГСЗ],
    конфиг: dict[str, Any],
    частоты: dict[str, int],
) -> КлючиСтроки:
    """Подбор минимального набора AND-ключей."""
    порог_not: int = int(конфиг["short_token_not_max_len"])
    кандидаты = отсортировать_токены(запись.токены, частоты)

    if not кандидаты:
        кандидаты = отсортировать_токены(
            извлечь_токены(запись.короткое_имя, int(конфиг["min_token_length"]), резервный=True),
            частоты,
        )

    if not кандидаты:
        return КлючиСтроки(дубликат=True, комментарий="нет токенов для ключа")

    def перебор(токен_список: list[str]) -> КлючиСтроки | None:
        варианты = [(т, выбрать_режим_токена(т, порог_not)) for т in токен_список]
        for размер in range(1, min(4, len(варианты) + 1)):
            for комбо in combinations(варианты, размер):
                ключи = построить_ключи_из_токенов(list(комбо))
                if уникально_для_записи(ключи, запись, все):
                    return ключи
        for размер in range(1, min(4, len(токен_список) + 1)):
            for индексы in combinations(range(len(токен_список)), размер):
                for маска in range(2 ** len(индексы)):
                    комбо = []
                    for бит, idx in enumerate(индексы):
                        is_full = bool(маска & (1 << бит))
                        комбо.append((токен_список[idx], is_full))
                    ключи = построить_ключи_из_токенов(комбо)
                    if уникально_для_записи(ключи, запись, все):
                        return ключи
        return None

    найдено = перебор(кандидаты)
    if найдено:
        return найдено

    # Один редкий токен (частота = 1) — уникален по определению в корпусе
    for т in кандидаты:
        if частоты.get(т.lower(), 0) == 1:
            ключи = построить_ключи_из_токенов([(т, выбрать_режим_токена(т, порог_not))])
            if строка_совпала(запись.короткое_имя, ключи):
                return ключи

    комбо = [(т, выбрать_режим_токена(т, порог_not)) for т in кандидаты[:3]]
    ключи = построить_ключи_из_токенов(комбо)
    ключи.дубликат = True
    ключи.комментарий = "не удалось подобрать уникальные ключи"
    return ключи


def прочитать_список(путь: Path) -> list[str]:
    """Чтение строк «Наименование, регион» из файла."""
    строки: list[str] = []
    with путь.open(encoding="utf-8") as файл:
        for номер, строка in enumerate(файл, start=1):
            строка = строка.strip()
            if номер <= 10 or not строка:
                continue
            if строка.startswith("#"):
                continue
            строки.append(строка)
    return строки


def отметить_групповые_дубликаты(записи: list[ЗаписьГСЗ]) -> None:
    """Пометка строк с одинаковыми ключами или пересекающимся совпадением."""
    for i, а in enumerate(записи):
        for б in записи[i + 1 :]:
            if а.ключи.в_словарь() == б.ключи.в_словарь() and а.ключи.в_словарь() != {к: "" for к in КОЛОНКИ_КЛЮЧЕЙ}:
                а.ключи.дубликат = True
                б.ключи.дубликат = True
                if not а.ключи.комментарий:
                    а.ключи.комментарий = "одинаковые ключи с другой строкой"
                if not б.ключи.комментарий:
                    б.ключи.комментарий = "одинаковые ключи с другой строкой"

    # Пересечение: один холдинг подходит к нескольким строкам
    for i, а in enumerate(записи):
        for б in записи[i + 1 :]:
            if строка_совпала(а.короткое_имя, б.ключи) and строка_совпала(а.короткое_имя, а.ключи):
                а.ключи.дубликат = True
                б.ключи.дубликат = True
            if строка_совпала(б.короткое_имя, а.ключи) and строка_совпала(б.короткое_имя, б.ключи):
                а.ключи.дубликат = True
                б.ключи.дубликат = True


def сохранить_excel(записи: list[ЗаписьГСЗ], путь: Path) -> None:
    """Экспорт в Excel со смарт-таблицей."""
    путь.parent.mkdir(parents=True, exist_ok=True)

    заголовки = ["Наименование, регион"] + КОЛОНКИ_КЛЮЧЕЙ + ["ключи_задублированы", "комментарий_ключей"]

    книга = Workbook()
    лист = книга.active
    лист.title = "_base_gsz"

    лист.append(заголовки)
    for запись in записи:
        словарь = запись.ключи.в_словарь()
        строка = [запись.полное_имя]
        строка.extend(словарь.get(к, "") for к in КОЛОНКИ_КЛЮЧЕЙ)
        строка.append("да" if запись.ключи.дубликат else "")
        строка.append(запись.ключи.комментарий)
        лист.append(строка)

    последняя_строка = len(записи) + 1
    последняя_колонка = chr(ord("A") + len(заголовки) - 1)
    ref = f"A1:{последняя_колонка}{последняя_строка}"

    таблица = Table(displayName="_base_gsz", ref=ref)
    таблица.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    лист.add_table(таблица)

    книга.save(путь)


def main() -> None:
    """Точка входа."""
    настроить_логирование()
    конфиг = загрузить_конфиг()

    вход = Path(конфиг["input_file"])
    выход = Path(конфиг["output_file"])
    if not вход.is_absolute():
        вход = КОРЕНЬ / вход
    if not выход.is_absolute():
        выход = КОРЕНЬ / выход
    мин_длина = int(конфиг["min_token_length"])

    logging.info("Чтение списка из %s", вход)
    полные_имена = прочитать_список(вход)
    logging.info("Загружено строк: %d", len(полные_имена))

    записи: list[ЗаписьГСЗ] = []
    for полное in полные_имена:
        короткое = полное.split(",", 1)[0].strip()
        токены = извлечь_токены(короткое, мин_длина)
        записи.append(ЗаписьГСЗ(полное_имя=полное, короткое_имя=короткое, токены=токены))

    logging.info("Подбор ключей...")
    частоты = частота_токенов(записи)
    for запись in записи:
        запись.ключи = подобрать_ключи(запись, записи, конфиг, частоты)

    отметить_групповые_дубликаты(записи)

    дубликатов = sum(1 for з in записи if з.ключи.дубликат)
    без_токенов = sum(1 for з in записи if not з.токены)

    сохранить_excel(записи, выход)

    logging.info("Сохранено: %s", выход)
    logging.info("Строк с пометкой дубликат: %d", дубликатов)
    logging.info("Строк без токенов: %d", без_токенов)


if __name__ == "__main__":
    main()
