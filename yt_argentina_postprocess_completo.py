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
SHEET_RANK_NUM1 = "Ranking Números 1"
SHEET_RANK_YT = "Ranking llista Completa"
SHEET_RANK_TOP10 = "Ranking primeres 10 cançons"

LLISTA_CONST = "YTCHar"
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


def build_ranking(rows: List[List[str]]) -> List[List[str]]:
    grouped = defaultdict(list)
    for r in rows:
        grouped[song_key(r[1], r[2])].append(r)

    out = []
    for _, grouped_rows in grouped.items():
        rows_sorted = sorted(grouped_rows, key=lambda x: parse_date(x[3]))
        out.append([
            rows_sorted[0][1],
            rows_sorted[0][2],
            str(len(rows_sorted)),
            rows_sorted[0][3],
            rows_sorted[-1][3],
            LLISTA_CONST,
            PAIS_CONST,
        ])

    out.sort(key=lambda x: (-int(x[2]), parse_date(x[3]), x[0].casefold(), x[1].casefold()))
    return out


def build_ranking_num1(source_rows: List[List[str]]) -> List[List[str]]:
    return build_ranking([r for r in source_rows if r[0] == "1"])


def build_ranking_full(source_rows: List[List[str]]) -> List[List[str]]:
    return build_ranking(source_rows)


def build_ranking_top10_without_num1(source_rows: List[List[str]]) -> List[List[str]]:
    num1_keys = {song_key(r[1], r[2]) for r in source_rows if r[0] == "1"}
    top10_candidate_keys = {
        song_key(r[1], r[2])
        for r in source_rows
        if r[0].isdigit() and 2 <= int(r[0]) <= 10
    }
    eligible_keys = top10_candidate_keys - num1_keys
    rows_for_ranking = [r for r in source_rows if song_key(r[1], r[2]) in eligible_keys]
    return build_ranking(rows_for_ranking)


def main():
    gc = get_gspread_client()
    spreadsheet = open_spreadsheet(gc)

    ws_yt = get_or_create_ws(spreadsheet, SHEET_YT, rows=200000, cols=10)
    ws_rank_num1 = get_or_create_ws(spreadsheet, SHEET_RANK_NUM1, rows=50000, cols=10)
    ws_rank_yt = get_or_create_ws(spreadsheet, SHEET_RANK_YT, rows=100000, cols=10)
    ws_rank_top10 = get_or_create_ws(spreadsheet, SHEET_RANK_TOP10, rows=50000, cols=10)

    ensure_headers(ws_yt, HEADERS_STANDARD)
    ensure_headers(ws_rank_num1, HEADERS_RANKING)
    ensure_headers(ws_rank_yt, HEADERS_RANKING)
    ensure_headers(ws_rank_top10, HEADERS_RANKING)

    yt_rows = standardize_source_rows(ws_yt.get_all_values()[1:])

    rank_num1_rows = build_ranking_num1(yt_rows)
    write_full_sheet(ws_rank_num1, HEADERS_RANKING, rank_num1_rows)
    print(f"Actualizada '{SHEET_RANK_NUM1}' con {len(rank_num1_rows)} filas.")

    rank_yt_rows = build_ranking_full(yt_rows)
    write_full_sheet(ws_rank_yt, HEADERS_RANKING, rank_yt_rows)
    print(f"Actualizada '{SHEET_RANK_YT}' con {len(rank_yt_rows)} filas.")

    rank_top10_rows = build_ranking_top10_without_num1(yt_rows)
    write_full_sheet(ws_rank_top10, HEADERS_RANKING, rank_top10_rows)
    print(f"Actualizada '{SHEET_RANK_TOP10}' con {len(rank_top10_rows)} filas.")


if __name__ == "__main__":
    main()

