#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Автоматический поиск УНП организаций и ИП (Беларусь) по названию.

Источник данных: открытый API Единого государственного регистра юридических
лиц и индивидуальных предпринимателей (ЕГР) — https://www.egr.gov.by/

Принцип работы
--------------
1. Читает список названий из xlsx/csv (столбец задаётся через --column).
2. Для каждого названия запрашивает ЕГР поиск по наименованию.
3. Нормализует названия (Ё→Е, регистр, лишние пробелы, кавычки, ООО/ЗАО/ИП).
4. Сопоставляет найденные варианты с исходным названием с помощью нечёткого
   сравнения (rapidfuzz) и выбирает лучший + считает балл уверенности.
5. Сохраняет промежуточные результаты в кэш (JSON) — повторный запуск
   продолжается с того же места, не опрашивая ЕГР повторно.
6. Выгружает итоговый Excel: исходное название, УНП, найденное название,
   статус в реестре, балл, решение (auto / review / not_found), кандидаты.

ВАЖНО: сервер egr.gov.by блокирует некоторые зарубежные/"датацентр" IP-адреса.
Если поиск не идёт, запускайте скрипт с компьютера, у которого есть доступ к
egr.gov.by (например, из белорусской сети). Используйте ключ --probe, чтобы
посмотреть "сырой" ответ сервера и при необходимости подкорректировать парсер.

Установка зависимостей:
    pip install requests openpyxl rapidfuzz

Запуск:
    python find_unp.py input.xlsx --column name --output result.xlsx
    python find_unp.py input.csv  --column 1            --output result.xlsx
    python find_unp.py input.xlsx --column name --probe     # отладка одного запроса
"""

import argparse
import csv
import json
import re
import sys
import time
from pathlib import Path
from urllib.parse import quote_plus

try:
    import requests
except ImportError:
    sys.exit("Установите requests:  pip install requests")

try:
    from rapidfuzz import fuzz
except ImportError:
    sys.exit("Установите rapidfuzz:  pip install rapidfuzz")

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
except ImportError:
    sys.exit("Установите openpyxl:  pip install openpyxl")


# --------------------------------------------------------------------------
# Конфигурация API ЕГР
# --------------------------------------------------------------------------
# Основной (и запасной) эндпоинты поиска по наименованию. Если основной
# недоступен/изменился — скрипт перебирает варианты по очереди.
EGR_API_ENDPOINTS = [
    "https://www.egr.gov.by/api/v2/ons/getRegisteredList?name={name}",
    "https://www.egr.gov.by/egr2api/api/v1/ons?searchType=name&searchValue={name}",
]
EGR_TIMEOUT = 25
EGR_DELAY = 0.4      # пауза между запросами, сек
EGR_RETRIES = 3

SCORE_ACCEPT = 92     # >= → auto
SCORE_REVIEW = 70     # >= и < ACCEPT → review


# --------------------------------------------------------------------------
# Нормализация названий
# --------------------------------------------------------------------------
_LEGAL_FORMS = [
    "ооо", "зоу", "зоо", "зао", "оао", "иу", "оу", "коу", "мп", "чуп", "чтуп",
    "ип", "общество с ограниченной ответственностью",
    "закрытое акционерное общество", "открытое акционерное общество",
    "частное унитарное предприятие", "коммунальное унитарное предприятие",
]
_QUOTE_RE = re.compile(r'[«»“”\'"`]')
_WS_RE = re.compile(r"\s+")


def normalize(name: str) -> str:
    if not name:
        return ""
    s = name.lower().strip()
    s = s.replace("ё", "е")
    s = _QUOTE_RE.sub(" ", s)
    s = _WS_RE.sub(" ", s).strip()
    for lf in _LEGAL_FORMS:
        s = re.sub(rf"\b{re.escape(lf)}\b", " ", s)
    s = re.sub(r"[^\w\s-]", " ", s)
    s = _WS_RE.sub(" ", s).strip()
    return s


# --------------------------------------------------------------------------
# Парсер ответа ЕГР (толерантный к схеме)
# --------------------------------------------------------------------------
_UNP_RE = re.compile(r"\b\d{9}\b")


def _find_unp_number(v):
    """Извлечь 9-значный УНП из значения любого типа (строка/число/список)."""
    if isinstance(v, list):
        for item in v:
            r = _find_unp_number(item)
            if r:
                return r
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, int):
        s = str(v)
        if len(s) == 9 and s.isdigit():
            return s
        m = _UNP_RE.search(s)
        return m.group() if m else None
    if isinstance(v, str):
        m = _UNP_RE.search(v)
        return m.group() if m else None
    return None


def _extract_candidates(obj):
    """Рекурсивно собрать из произвольного JSON список словарей-кандидатов.

    Каждый кандидат — dict, содержащий хотя бы одно поле-УНП (9 цифр) и
    хотя бы одно строковое поле с наименованием. list[dict].
    """
    candidates = []

    def walk(node):
        if isinstance(node, dict):
            unp = None
            name = None
            # 1) УНП в полях с «ключевым» именем
            for k, v in node.items():
                ks = str(k).lower()
                if any(t in ks for t in ("unp", "vnum", "inu", "inn", "regnum",
                                         "reg_num", "nomer", "номер", "учетн",
                                         "учн", "ogr", "payer")):
                    r = _find_unp_number(v)
                    if r and unp is None:
                        unp = r
            # 2) иначе — любое поле со значением, похожим на УНП
            if unp is None:
                for v in node.values():
                    r = _find_unp_number(v)
                    if r:
                        unp = r
                        break
            def _first_str(v):
                """Достать первую непустую строку (в т.ч. из списка)."""
                if isinstance(v, str):
                    return v.strip() if v.strip() else None
                if isinstance(v, list):
                    for it in v:
                        r = _first_str(it)
                        if r:
                            return r
                return None
            # 3) наименование по ключевому имени поля
            for k, v in node.items():
                ks = str(k).lower()
                if any(t in ks for t in ("name", "naim", "vnaim", "naimp",
                                         "naimov", "polnoe", "полное",
                                         "наименов", "fulln", "firm", "nazv")):
                    r = _first_str(v)
                    if r and len(r) > 2:
                        name = r
                        break
            # 4) иначе — самая длинная строка в этом dict (в т.ч. внутри списков)
            if name is None:
                all_strs = []
                for v in node.values():
                    if isinstance(v, str) and len(v) > 3 and not v.isdigit():
                        all_strs.append(v.strip())
                    elif isinstance(v, list):
                        for it in v:
                            if isinstance(it, str) and len(it) > 3 and not it.isdigit():
                                all_strs.append(it.strip())
                if all_strs:
                    name = max(all_strs, key=len).strip()
            if unp and name:
                status = None
                for k, v in node.items():
                    ks = str(k).lower()
                    if isinstance(v, str) and any(t in ks for t in
                                                 ("status", "sost", "состоян", "сост")):
                        status = v.strip()
                        break
                # не используем в качестве name технические поля
                if name.lower() not in ("действует", "ликвидирован", "активно"):
                    candidates.append({"unp": unp, "name": name,
                                       "status": status, "raw": node})
            for v in node.values():
                walk(v)
        elif isinstance(node, list):
            for item in node:
                walk(item)

    walk(obj)
    return candidates


# --------------------------------------------------------------------------
# Поиск в ЕГР
# --------------------------------------------------------------------------
_session = requests.Session()
_session.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "ru-RU,ru;q=0.9",
})


def egr_search(name: str, debug=False):
    """Вернуть список кандидатов (dict) из ЕГР по названию."""
    enc = quote_plus(name)
    last_text = None
    for url_tmpl in EGR_API_ENDPOINTS:
        url = url_tmpl.format(name=enc)
        for attempt in range(1, EGR_RETRIES + 1):
            try:
                r = _session.get(url, timeout=EGR_TIMEOUT)
                if r.status_code == 200:
                    try:
                        data = r.json()
                    except ValueError:
                        last_text = r.text[:1000]
                        break
                    if debug:
                        print(f"\n[probe] {url}\nstatus=200\nraw:\n"
                              f"{json.dumps(data, ensure_ascii=False)[:2000]}\n",
                              flush=True)
                    cands = _extract_candidates(data)
                    if cands:
                        return cands
                    last_text = json.dumps(data, ensure_ascii=False)[:1000]
                    break
                else:
                    last_text = f"HTTP {r.status_code}: {r.text[:300]}"
                    if r.status_code == 404:
                        break
            except requests.RequestException as e:
                last_text = str(e)
                time.sleep(1.0 * attempt)
    if debug:
        print(f"[probe] кандидаты не найдены. last={last_text}", flush=True)
    return []


# --------------------------------------------------------------------------
# Сопоставление
# --------------------------------------------------------------------------
def best_match(input_name: str, candidates):
    norm_in = normalize(input_name)
    best = None
    best_score = -1.0
    for c in candidates:
        score = fuzz.token_sort_ratio(norm_in, normalize(c["name"]))
        if score > best_score:
            best_score = score
            best = c
    if best is None:
        return None, 0.0, []
    ranked = sorted(
        [{"unp": c["unp"], "name": c["name"],
          "score": round(fuzz.token_sort_ratio(norm_in, normalize(c["name"])), 1)}
         for c in candidates],
        key=lambda x: x["score"], reverse=True,
    )[:5]
    return best, round(best_score, 1), ranked


def decide(score: float) -> str:
    if score >= SCORE_ACCEPT:
        return "auto"
    if score >= SCORE_REVIEW:
        return "review"
    return "low"


# --------------------------------------------------------------------------
# Чтение списка
# --------------------------------------------------------------------------
def _resolve_column(header, column):
    if header is None:
        return max(0, int(column) - 1)
    if isinstance(column, str) and not column.isdigit():
        for i, h in enumerate(header):
            if h is not None and str(h).strip().lower() == column.strip().lower():
                return i
        raise SystemExit(f"Столбец «{column}» не найден. Заголовки: {header}")
    return int(column) - 1


def read_names(path: str, column):
    p = Path(path)
    names = []
    if p.suffix.lower() in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        ws = wb.active
        rows = ws.iter_rows(values_only=True)
        header = next(rows, None)
        col_idx = _resolve_column(header, column)
        for i, row in enumerate(rows, start=2):
            val = row[col_idx] if col_idx < len(row) else None
            if val is not None and str(val).strip():
                names.append((i, str(val).strip()))
    else:
        with open(p, newline="", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            header = next(reader, None)
            col_idx = _resolve_column(header, column)
            for i, row in enumerate(reader, start=2):
                val = row[col_idx] if col_idx < len(row) else None
                if val and val.strip():
                    names.append((i, val.strip()))
    return names


# --------------------------------------------------------------------------
# Выгрузка в Excel
# --------------------------------------------------------------------------
_HEADER_FILL = PatternFill("solid", fgColor="305496")
_HEADER_FONT = Font(color="FFFFFF", bold=True)
_AUTO_FILL = PatternFill("solid", fgColor="C6EFCE")
_REVIEW_FILL = PatternFill("solid", fgColor="FFEB9C")
_LOW_FILL = PatternFill("solid", fgColor="FFC7CE")


def write_excel(rows, out_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "УНП"
    headers = ["№ строки", "Исходное название", "УНП", "Найденное название",
               "Статус", "Балл", "Решение", "Кандидаты (УНП | название | балл)"]
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for r in rows:
        ws.append(r)
        fill = {"auto": _AUTO_FILL, "review": _REVIEW_FILL,
                "low": _LOW_FILL, "not_found": _LOW_FILL}.get(r[6], None)
        if fill:
            for c in range(1, len(headers) + 1):
                ws.cell(row=ws.max_row, column=c).fill = fill
    widths = [8, 52, 12, 52, 16, 8, 10, 60]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    ws2 = wb.create_sheet("Сводка")
    from collections import Counter
    cnt = Counter(r[6] for r in rows)
    ws2.append(["Решение", "Кол-во", "Доля"])
    for c in range(1, 4):
        ws2.cell(row=1, column=c).fill = _HEADER_FILL
        ws2.cell(row=1, column=c).font = _HEADER_FONT
    total = max(1, len(rows))
    for k in ("auto", "review", "low", "not_found"):
        if cnt.get(k):
            ws2.append([k, cnt[k], f"{cnt[k] / total:.1%}"])
    ws2.append(["ВСЕГО", len(rows), "100%"])
    wb.save(out_path)


# --------------------------------------------------------------------------
# Кэш
# --------------------------------------------------------------------------
def load_cache(path):
    if Path(path).exists():
        try:
            with open(path, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_cache(path, cache):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=0)
    Path(tmp).replace(path)


# --------------------------------------------------------------------------
# Главная
# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(
        description="Поиск УНП (Беларусь) по названию через ЕГР")
    ap.add_argument("input", help="Файл со списком названий (xlsx/csv)")
    ap.add_argument("--column", default="1",
                    help="Имя столбца или его номер (1-индекс). По умолч. 1")
    ap.add_argument("--output", default="unp_result.xlsx",
                    help="Файл результата (xlsx)")
    ap.add_argument("--cache", default="unp_cache.json",
                    help="Файл кэша для возобновляемого поиска")
    ap.add_argument("--probe", action="store_true",
                    help="Отладка: вывести сырой ответ для первой записи")
    args = ap.parse_args()

    names = read_names(args.input, args.column)
    if not names:
        sys.exit("Список названий пуст — проверьте --column и содержимое файла.")
    print(f"Загружено названий: {len(names)}")

    cache = load_cache(args.cache)

    if args.probe:
        sample = names[0][1]
        print(f"[probe] Тестируем запрос для: {sample!r}", flush=True)
        egr_search(sample, debug=True)
        print('[probe] Готово. Если кандидаты пусты — поправьте '
              'EGR_API_ENDPOINTS/парсер по сырому ответу.')
        return

    out_rows = []
    for idx, (row_no, name) in enumerate(names, start=1):
        result = cache.get(name)
        if result is None:
            try:
                cands = egr_search(name)
            except Exception as e:
                cands = []
                print(f"  ! ошибка запроса для «{name}»: {e}")
            best, score, ranked = (best_match(name, cands) if cands
                                   else (None, 0.0, []))
            decision = decide(score) if best else "not_found"
            result = {
                "unp": best["unp"] if best else "",
                "matched": best["name"] if best else "",
                "status": (best["status"] if best and best.get("status") else ""),
                "score": score,
                "decision": decision,
                "candidates": ranked,
            }
            cache[name] = result
            if idx % 10 == 0 or idx == len(names):
                save_cache(args.cache, cache)
            time.sleep(EGR_DELAY)

        cands_str = " ; ".join(
            f"{c['unp']} | {c['name']} | {c['score']}"
            for c in result.get("candidates", []))
        out_rows.append([
            row_no, name, result["unp"], result["matched"],
            result["status"], result["score"], result["decision"], cands_str,
        ])
        bar = ("OK " if result["decision"] == "auto"
               else ("?  " if result["decision"] == "review" else "-- "))
        print(f"{bar}{idx}/{len(names)}  {result['decision']:>9}  "
              f"{result['score']:>5}  {result['unp'] or '—':<9}  {name[:40]}")

    save_cache(args.cache, cache)
    write_excel(out_rows, args.output)
    found = sum(1 for r in out_rows if r[6] in ("auto", "review"))
    print(f"\nГотово. Найдено (auto+review): {found}/{len(out_rows)}  →  {args.output}")


if __name__ == "__main__":
    main()

