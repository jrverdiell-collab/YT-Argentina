#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from typing import List, Tuple

import gspread
from gspread.exceptions import WorksheetNotFound, APIError
from google.oauth2.service_account import Credentials

SERVICE_ACCOUNT_FILE = "service_account.json"
SPREADSHEET_ID = "1FoL3lZNvIaC49wC6O_TKp3PNk3fS7YI3FSnZkf1U1iQ"

SHEET_YT = "YT Argentina"
SHEET_NEW = "Cançons noves"
SHEET_TOP10 = "Primeres 10 cançons"
SHEET_NUM1 = "Números 1"
SHEET_NUM1_DEP = "Num. 1 Depurats"
SHEET_RANK_NUM1 = "Ranking Números 1"
SHEET_RANK_YT = "Ranking llista completa"

LLISTA_CONST = "YTCHArg"
PAIS_CONST = "Argentina"

HEADERS_STANDARD = ["Núm. Lista", "Cançó", "Interpret", "Data", "Llista", "Pais"]
HEADERS_RANKING = ["Cançó", "Interpret", "Núm. Setmanes", "Primera Data", "Ultima Data", "Llista", "Pais"]


def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
    return gspread.authorize(creds)


def open_spreadsheet(gc: gspread.Client):
    return gc.open_by_key(SPREADSHEET_ID)


def normalize_title(s: str) -> str:
    return " ".join((s or "").split()).strip().casefold()


def get_or_create_ws(spreadsheet, title: str, rows: int = 1000, cols: int = 10):
    wanted = normalize_title(title)

    try:
        return spreadsheet.worksheet(title)
    except WorksheetNotFound:
        pass

    for ws in spreadsheet.worksheets():
        if normalize_title(ws.title) == wanted:
            return ws

    try:
        return spreadsheet.add_worksheet(title=title, rows=rows, cols=cols)
    except APIError as e:
        msg = str(e)
        if "already exists" in msg or "ya existe" in msg:
            refreshed = spreadsheet.client.open_by_key(spreadsheet.id)
            for ws in refreshed.worksheets():
                if normalize_title(ws.title) == wanted:
                    return ws
        raise


def ensure_headers(ws, headers):
    current = ws.row_values(1)
    if current != headers:
        ws.clear()
        ws.update(range_name="A1", values=[headers])


def parse_date(s: str) -> datetime:
    return datetime.strptime(s.strip(), "%d/%m/%Y")


def normalize_text(s: str) -> str:
    return " ".join((s or "").split()).strip()


def song_key(song: str, artist: str) -> Tuple[str, str]:
    return (normalize_text(song).casefold(), normalize_text(artist).casefold())


def full_key(row: List[str]) -> Tuple[str, str, str]:
    return (
        normalize_text(row[1]).casefold(),
        normalize_text(row[2]).casefold(),
        normalize_text(row[3]),
    )


def standardize_source_rows(raw_rows: List[List[str]]) -> List[List[str]]:
    out = []
    for r in raw_rows:
        if len(r) < 6:
            continue
        rank = normalize_text(r[0])
        song = normalize_text(r[1])
        artist = normalize_text(r[2])
        date = normalize_text(r[3])
        if rank.isdigit() and song and date:
            out.append([rank, song, artist, date, LLISTA_CONST, PAIS_CONST])
    return out


def write_full_sheet(ws, headers, rows):
    ws.clear()
    ws.update(range_name="A1", values=[headers] + rows)


def build_cancons_noves(source_rows: List[List[str]], existing_rows: List[List[str]]) -> List[List[str]]:
    existing_song_keys = {song_key(r[1], r[2]) for r in existing_rows if len(r) >= 6}
    out = existing_rows[:]

    for r in sorted(source_rows, key=lambda x: parse_date(x[3])):
        k = song_key(r[1], r[2])
        if k not in existing_song_keys:
            out.append(r)
            existing_song_keys.add(k)

    return out


def build_primeres_10(source_rows: List[List[str]], existing_rows: List[List[str]]) -> List[List[str]]:
    current_map = {}

    for r in existing_rows:
        if len(r) >= 6:
            current_map[song_key(r[1], r[2])] = r

    num1_keys = {song_key(r[1], r[2]) for r in source_rows if r[0] == "1"}

    for k in list(current_map.keys()):
        if k in num1_keys:
            del current_map[k]

    source_sorted = sorted(source_rows, key=lambda x: (parse_date(x[3]), int(x[0])))

    for r in source_sorted:
        rank = int(r[0])
        if 2 <= rank <= 10:
            k = song_key(r[1], r[2])
            if k not in current_map:
                current_map[k] = r

    out = list(current_map.values())
    out.sort(key=lambda x: (parse_date(x[3]), int(x[0]), x[1].casefold(), x[2].casefold()))
    return out


def build_numeros_1(source_rows: List[List[str]], existing_rows: List[List[str]]) -> List[List[str]]:
    existing_keys = {full_key(r) for r in existing_rows if len(r) >= 6}
    out = existing_rows[:]

    for r in sorted([x for x in source_rows if x[0] == "1"], key=lambda x: parse_date(x[3])):
        k = full_key(r)
        if k not in existing_keys:
            out.append(r)
            existing_keys.add(k)

    return out


def build_num1_depurats(num1_rows: List[List[str]], existing_rows: List[List[str]]) -> List[List[str]]:
    existing_song_keys = {song_key(r[1], r[2]) for r in existing_rows if len(r) >= 6}
    out = existing_rows[:]

    for r in sorted(num1_rows, key=lambda x: parse_date(x[3])):
        k = song_key(r[1], r[2])
        if k not in existing_song_keys:
            out.append(r)
            existing_song_keys.add(k)

    return out


def build_ranking(source_rows: List[List[str]]) -> List[List[str]]:
    grouped = defaultdict(list)

    for r in source_rows:
        grouped[song_key(r[1], r[2])].append(r)

    out = []

    for (_song_key, _artist_key), rows in grouped.items():
        rows_sorted = sorted(rows, key=lambda x: parse_date(x[3]))
        song = rows_sorted[0][1]
        artist = rows_sorted[0][2]
        num_setmanes = len(rows_sorted)
        primera_data = rows_sorted[0][3]
        ultima_data = rows_sorted[-1][3]

        out.append([
            song,
            artist,
            str(num_setmanes),
            primera_data,
            ultima_data,
            LLISTA_CONST,
            PAIS_CONST
        ])

    out.sort(key=lambda x: (-int(x[2]), parse_date(x[3]), x[0].casefold(), x[1].casefold()))
    return out


def main():
    gc = get_gspread_client()
    spreadsheet = open_spreadsheet(gc)

    ws_yt = get_or_create_ws(spreadsheet, SHEET_YT, rows=200000, cols=10)
    ws_new = get_or_create_ws(spreadsheet, SHEET_NEW, rows=50000, cols=10)
    ws_top10 = get_or_create_ws(spreadsheet, SHEET_TOP10, rows=50000, cols=10)
    ws_num1 = get_or_create_ws(spreadsheet, SHEET_NUM1, rows=50000, cols=10)
    ws_num1_dep = get_or_create_ws(spreadsheet, SHEET_NUM1_DEP, rows=50000, cols=10)
    ws_rank_num1 = get_or_create_ws(spreadsheet, SHEET_RANK_NUM1, rows=50000, cols=10)
    ws_rank_yt = get_or_create_ws(spreadsheet, SHEET_RANK_YT, rows=100000, cols=10)

    ensure_headers(ws_yt, HEADERS_STANDARD)
    ensure_headers(ws_new, HEADERS_STANDARD)
    ensure_headers(ws_top10, HEADERS_STANDARD)
    ensure_headers(ws_num1, HEADERS_STANDARD)
    ensure_headers(ws_num1_dep, HEADERS_STANDARD)
    ensure_headers(ws_rank_num1, HEADERS_RANKING)
    ensure_headers(ws_rank_yt, HEADERS_RANKING)

    yt_rows = standardize_source_rows(ws_yt.get_all_values()[1:])
    new_rows_existing = standardize_source_rows(ws_new.get_all_values()[1:])
    top10_rows_existing = standardize_source_rows(ws_top10.get_all_values()[1:])
    num1_rows_existing = standardize_source_rows(ws_num1.get_all_values()[1:])
    num1_dep_rows_existing = standardize_source_rows(ws_num1_dep.get_all_values()[1:])

    new_rows_final = build_cancons_noves(yt_rows, new_rows_existing)
    write_full_sheet(ws_new, HEADERS_STANDARD, new_rows_final)
    print(f"Actualizada '{SHEET_NEW}' con {len(new_rows_final)} filas.")

    top10_rows_final = build_primeres_10(yt_rows, top10_rows_existing)
    write_full_sheet(ws_top10, HEADERS_STANDARD, top10_rows_final)
    print(f"Actualizada '{SHEET_TOP10}' con {len(top10_rows_final)} filas.")

    num1_rows_final = build_numeros_1(yt_rows, num1_rows_existing)
    write_full_sheet(ws_num1, HEADERS_STANDARD, num1_rows_final)
    print(f"Actualizada '{SHEET_NUM1}' con {len(num1_rows_final)} filas.")

    num1_dep_rows_final = build_num1_depurats(num1_rows_final, num1_dep_rows_existing)
    write_full_sheet(ws_num1_dep, HEADERS_STANDARD, num1_dep_rows_final)
    print(f"Actualizada '{SHEET_NUM1_DEP}' con {len(num1_dep_rows_final)} filas.")

    rank_num1_rows = build_ranking(num1_rows_final)
    write_full_sheet(ws_rank_num1, HEADERS_RANKING, rank_num1_rows)
    print(f"Actualizada '{SHEET_RANK_NUM1}' con {len(rank_num1_rows)} filas.")

    rank_yt_rows = build_ranking(yt_rows)
    write_full_sheet(ws_rank_yt, HEADERS_RANKING, rank_yt_rows)
    print(f"Actualizada '{SHEET_RANK_YT}' con {len(rank_yt_rows)} filas.")


if __name__ == "__main__":
    main()