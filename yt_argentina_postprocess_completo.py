#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

from collections import defaultdict
from datetime import datetime

import gspread
from google.oauth2.service_account import Credentials


SERVICE_ACCOUNT_FILE = "service_account.json"
SPREADSHEET_ID = "1FoL3lZNvIaC49wC6O_TKp3PNk3fS7YI3FSnZkf1U1iQ"

SHEET_YT = "YT Argentina"
SHEET_RANK_NUM1 = "Ranking Números 1"
SHEET_RANK_YT = "Ranking llista Completa"
SHEET_RANK_TOP10 = "Ranking primeres 10 cançons"

LLISTA_CONST = "YTCHArg"
PAIS_CONST = "Argentina"

HEADERS_RANKING = [
    "Cançó",
    "Interpret",
    "Núm. Setmanes",
    "Primera Data",
    "Ultima Data",
    "Millor posició",
    "Llista",
    "Pais",
]

HEADERS_RANKING_NUM1 = [
    "Cançó",
    "Interpret",
    "Núm. Setmanes",
    "Primera Data",
    "Ultima Data",
    "Llista",
    "Pais",
]


def get_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
    return gspread.authorize(creds)


def normalize_text(s: str) -> str:
    return " ".join((s or "").split()).strip()


def parse_date(s: str) -> datetime:
    return datetime.strptime(s.strip(), "%d/%m/%Y")


def song_key(song: str, artist: str):
    return (
        normalize_text(song).casefold(),
        normalize_text(artist).casefold(),
    )


def row_key_same_day(row):
    return (
        normalize_text(row[1]).casefold(),
        normalize_text(row[2]).casefold(),
        normalize_text(row[3]),
    )


def standardize_source_rows(raw_rows):
    out = []

    for r in raw_rows:
        if len(r) < 6:
            continue

        rank = normalize_text(r[0])
        song = normalize_text(r[1])
        artist = normalize_text(r[2])
        date_str = normalize_text(r[3])

        if not rank.isdigit():
            continue
        if not song or not date_str:
            continue

        out.append([
            rank,
            song,
            artist,
            date_str,
            LLISTA_CONST,
            PAIS_CONST,
        ])

    return out


def dedupe_keep_best_rank(rows):
    best = {}

    for r in rows:
        key = row_key_same_day(r)
        rank = int(r[0])

        if key not in best or rank < int(best[key][0]):
            best[key] = r

    out = list(best.values())
    out.sort(key=lambda x: (parse_date(x[3]), int(x[0]), x[1].casefold(), x[2].casefold()))
    return out


def ensure_headers_only_if_missing(ws, headers):
    current = ws.row_values(1)
    if current != headers:
        ws.update(range_name="A1", values=[headers])


def clear_data_keep_header(ws):
    ws.batch_clear(["A2:H100000"])


def write_rows_keep_header(ws, rows):
    if rows:
        ws.update(range_name="A2", values=rows, value_input_option="USER_ENTERED")


def write_ranking_sheet(ws, rows):
    ensure_headers_only_if_missing(ws, HEADERS_RANKING)
    clear_data_keep_header(ws)
    write_rows_keep_header(ws, rows)


def write_ranking_num1_sheet(ws, rows):
    ensure_headers_only_if_missing(ws, HEADERS_RANKING_NUM1)
    ws.batch_clear(["A2:G100000"])
    if rows:
        ws.update(range_name="A2", values=rows, value_input_option="USER_ENTERED")


def apply_ranking_sheet_format_and_filter(spreadsheet, ws):
    """
    C = Núm. Setmanes -> número
    D = Primera Data -> fecha
    E = Ultima Data -> fecha
    F = Millor posició -> número
    """
    try:
        requests = [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": ws.id,
                        "startRowIndex": 1,
                        "endRowIndex": ws.row_count,
                        "startColumnIndex": 2,
                        "endColumnIndex": 3
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": "NUMBER",
                                "pattern": "0"
                            }
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat"
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": ws.id,
                        "startRowIndex": 1,
                        "endRowIndex": ws.row_count,
                        "startColumnIndex": 3,
                        "endColumnIndex": 5
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": "DATE",
                                "pattern": "dd/mm/yyyy"
                            }
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat"
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": ws.id,
                        "startRowIndex": 1,
                        "endRowIndex": ws.row_count,
                        "startColumnIndex": 5,
                        "endColumnIndex": 6
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": "NUMBER",
                                "pattern": "0"
                            }
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat"
                }
            },
            {
                "setBasicFilter": {
                    "filter": {
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": 0,
                            "endRowIndex": ws.row_count,
                            "startColumnIndex": 0,
                            "endColumnIndex": 8
                        }
                    }
                }
            }
        ]
        spreadsheet.batch_update({"requests": requests})
        print(f"Formato y filtro aplicados en '{ws.title}'.")
    except Exception as e:
        print(f"No se pudo aplicar formato/filtro en '{ws.title}': {e}")


def apply_ranking_num1_sheet_format_and_filter(spreadsheet, ws):
    """
    C = Núm. Setmanes -> número
    D = Primera Data -> fecha
    E = Ultima Data -> fecha
    """
    try:
        requests = [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": ws.id,
                        "startRowIndex": 1,
                        "endRowIndex": ws.row_count,
                        "startColumnIndex": 2,
                        "endColumnIndex": 3
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": "NUMBER",
                                "pattern": "0"
                            }
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat"
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": ws.id,
                        "startRowIndex": 1,
                        "endRowIndex": ws.row_count,
                        "startColumnIndex": 3,
                        "endColumnIndex": 5
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": "DATE",
                                "pattern": "dd/mm/yyyy"
                            }
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat"
                }
            },
            {
                "setBasicFilter": {
                    "filter": {
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": 0,
                            "endRowIndex": ws.row_count,
                            "startColumnIndex": 0,
                            "endColumnIndex": 7
                        }
                    }
                }
            }
        ]
        spreadsheet.batch_update({"requests": requests})
        print(f"Formato y filtro aplicados en '{ws.title}'.")
    except Exception as e:
        print(f"No se pudo aplicar formato/filtro en '{ws.title}': {e}")


def build_ranking(rows):
    grouped = defaultdict(list)

    for r in rows:
        grouped[song_key(r[1], r[2])].append(r)

    out = []

    for _, items in grouped.items():
        items_sorted = sorted(items, key=lambda x: parse_date(x[3]))
        song = items_sorted[0][1]
        artist = items_sorted[0][2]
        num_setmanes = len(items_sorted)
        primera_data = items_sorted[0][3]
        ultima_data = items_sorted[-1][3]
        millor_posicio = min(int(x[0]) for x in items_sorted)

        out.append([
            song,
            artist,
            str(num_setmanes),
            primera_data,
            ultima_data,
            str(millor_posicio),
            LLISTA_CONST,
            PAIS_CONST,
        ])

    out.sort(key=lambda x: (-int(x[2]), parse_date(x[3]), x[0].casefold(), x[1].casefold()))
    return out


def build_ranking_num1(rows):
    grouped = defaultdict(list)

    for r in rows:
        grouped[song_key(r[1], r[2])].append(r)

    out = []

    for _, items in grouped.items():
        items_sorted = sorted(items, key=lambda x: parse_date(x[3]))
        song = items_sorted[0][1]
        artist = items_sorted[0][2]
        num_setmanes = len(items_sorted)
        primera_data = items_sorted[0][3]
        ultima_data = items_sorted[-1][3]

        out.append([
            song,
            artist,
            str(num_setmanes),
            primera_data,
            ultima_data,
            LLISTA_CONST,
            PAIS_CONST,
        ])

    out.sort(key=lambda x: (-int(x[2]), parse_date(x[3]), x[0].casefold(), x[1].casefold()))
    return out


def build_ranking_top10(rows):
    songs_top10 = {
        song_key(r[1], r[2])
        for r in rows
        if 1 <= int(r[0]) <= 10
    }

    songs_num1 = {
        song_key(r[1], r[2])
        for r in rows
        if int(r[0]) == 1
    }

    valid_songs = songs_top10 - songs_num1

    filtered = [
        r for r in rows
        if song_key(r[1], r[2]) in valid_songs and 1 <= int(r[0]) <= 10
    ]

    return build_ranking(filtered)


def main():
    gc = get_client()
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)

    print("SPREADSHEET_ID =", SPREADSHEET_ID)
    print("SHEET_RANK_TOP10 =", repr(SHEET_RANK_TOP10))
    print("PESTAÑAS DISPONIBLES =", [ws.title for ws in spreadsheet.worksheets()])

    ws_yt = spreadsheet.worksheet(SHEET_YT)
    ws_rank_num1 = spreadsheet.worksheet(SHEET_RANK_NUM1)
    ws_rank_yt = spreadsheet.worksheet(SHEET_RANK_YT)
    ws_rank_top10 = spreadsheet.worksheet(SHEET_RANK_TOP10)

    yt_raw = ws_yt.get_all_values()[1:]
    yt_rows = standardize_source_rows(yt_raw)
    yt_rows = dedupe_keep_best_rank(yt_rows)

    rank_all = build_ranking(yt_rows)
    write_ranking_sheet(ws_rank_yt, rank_all)
    apply_ranking_sheet_format_and_filter(spreadsheet, ws_rank_yt)
    print(f"Actualizada '{SHEET_RANK_YT}' con {len(rank_all)} filas.")

    num1_rows = [r for r in yt_rows if r[0] == "1"]
    rank_num1 = build_ranking_num1(num1_rows)
    write_ranking_num1_sheet(ws_rank_num1, rank_num1)
    apply_ranking_num1_sheet_format_and_filter(spreadsheet, ws_rank_num1)
    print(f"Actualizada '{SHEET_RANK_NUM1}' con {len(rank_num1)} filas.")

    rank_top10 = build_ranking_top10(yt_rows)
    write_ranking_sheet(ws_rank_top10, rank_top10)
    apply_ranking_sheet_format_and_filter(spreadsheet, ws_rank_top10)
    print(f"Actualizada '{SHEET_RANK_TOP10}' con {len(rank_top10)} filas.")


if __name__ == "__main__":
    main()
