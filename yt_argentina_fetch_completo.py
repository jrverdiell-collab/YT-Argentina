#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
yt_argentina_fetch_completo.py
"""

from __future__ import annotations

import os
import re
import sys
import time
import subprocess
from datetime import datetime, timedelta

import gspread
from google.oauth2.service_account import Credentials
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# =========================================================
# CONFIG
# =========================================================

SERVICE_ACCOUNT_FILE = "service_account.json"

SPREADSHEET_ID = "1FoL3lZNvIaC49wC6O_TKp3PNk3fS7YI3FSnZkf1U1iQ"

SHEET_FULL = "Full 1"
SHEET_YT = "YT Argentina"

POSTPROCESS_SCRIPT = "yt_argentina_postprocess_completo.py"

LLISTA_CONST = "YTCHArg"
PAIS_CONST = "Argentina"

HEADERS_YT = ["Núm. Lista", "Cançó", "Interpret", "Data", "Llista", "Pais"]


# =========================================================
# FORMATO TEXTO
# =========================================================

ABBREVIATIONS = {
    "dj", "usa", "uk", "vol", "pt", "feat", "ft"
}

LOWERCASE_WORDS = {
    "x", "&", "feat.", "ft."
}


def smart_title_case(text: str) -> str:
    text = " ".join(text.split()).strip()

    words = text.split()
    formatted_words = []

    for w in words:
        clean = w.lower()

        if clean in LOWERCASE_WORDS:
            formatted_words.append(clean)
        elif clean.replace(".", "") in ABBREVIATIONS:
            formatted_words.append(clean.upper())
        else:
            formatted_words.append(clean.capitalize())

    return " ".join(formatted_words)


def format_song_title(text: str) -> str:
    text = " ".join(text.split()).strip()

    if not text:
        return text

    words = text.lower().split()
    words[0] = words[0].capitalize()

    return " ".join(words)


# =========================================================
# GOOGLE SHEETS
# =========================================================

def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    creds = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=scopes,
    )

    return gspread.authorize(creds)


def open_spreadsheet(gc):
    return gc.open_by_key(SPREADSHEET_ID)


def get_or_create_ws(spreadsheet, title, rows=1000, cols=10):
    try:
        return spreadsheet.worksheet(title)
    except gspread.WorksheetNotFound:
        return spreadsheet.add_worksheet(title=title, rows=rows, cols=cols)


def ensure_headers(ws, headers):
    current = ws.row_values(1)
    if current != headers:
        ws.update(range_name="A1", values=[headers])


def read_full1_a1(spreadsheet):
    ws = spreadsheet.worksheet(SHEET_FULL)
    return (ws.acell("A1").value or "").strip()


def write_full1_a1(spreadsheet, value):
    ws = spreadsheet.worksheet(SHEET_FULL)
    ws.update(range_name="A1", values=[[value]])


# =========================================================
# SELENIUM
# =========================================================

def build_driver():
    options = Options()

    options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--lang=es-ES")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")

    # En Windows no pongas rutas Linux como /usr/bin/...
    # webdriver-manager descargará y usará el chromedriver correcto.
    service = Service(ChromeDriverManager().install())

    driver = webdriver.Chrome(
        service=service,
        options=options
    )

    driver.set_page_load_timeout(90)

    return driver


# =========================================================
# URL FECHAS
# =========================================================

def extract_yyyymmdd_from_url(url):
    m = re.search(r"/weekly/(\d{8})", url)
    return m.group(1) if m else None


def next_week_url(url):
    date_str = extract_yyyymmdd_from_url(url)

    if not date_str:
        raise ValueError(f"No se pudo extraer la fecha de la URL: {url}")

    current_date = datetime.strptime(date_str, "%Y%m%d")
    next_date = current_date + timedelta(days=7)

    return url.replace(date_str, next_date.strftime("%Y%m%d"))


def format_date_from_url(url):
    date_str = extract_yyyymmdd_from_url(url)

    if not date_str:
        raise ValueError(f"No se pudo extraer la fecha de la URL: {url}")

    d = datetime.strptime(date_str, "%Y%m%d")
    return d.strftime("%d/%m/%Y")


# =========================================================
# EXTRACCION
# =========================================================

def accept_cookies_if_present(driver):
    xpaths = [
        "//button[contains(., 'Accept')]",
        "//button[contains(., 'Aceptar')]",
    ]

    for xp in xpaths:
        try:
            btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, xp))
            )
            btn.click()
            time.sleep(2)
            return
        except Exception:
            pass


def extract_rows(page_html, chart_date):
    soup = BeautifulSoup(page_html, "html.parser")

    rows = []
    seen = set()

    for row in soup.select("ytmc-entry-row"):
        rank_el = row.select_one("span#rank")
        title_el = row.select_one("div#entity-title")
        artist_els = row.select("span.artistName")

        if not rank_el or not title_el:
            continue

        rank = rank_el.text.strip()
        title = format_song_title(title_el.text.strip())

        artists = []
        for a in artist_els:
            t = smart_title_case(a.text.strip())
            if t and t not in artists:
                artists.append(t)

        artist = " & ".join(artists)

        if rank in seen:
            continue

        seen.add(rank)

        rows.append([
            rank,
            title,
            artist,
            chart_date,
            LLISTA_CONST,
            PAIS_CONST,
        ])

        if len(rows) >= 100:
            break

    rows.sort(key=lambda x: int(x[0]))
    return rows


def append_if_new(ws, rows):
    existing = ws.get_all_values()

    keys = {
        (
            r[0],
            r[1].lower(),
            r[2].lower(),
            r[3],
        )
        for r in existing[1:]
        if len(r) >= 4
    }

    new = []

    for r in rows:
        k = (
            r[0],
            r[1].lower(),
            r[2].lower(),
            r[3],
        )
        if k not in keys:
            new.append(r)

    if new:
        ws.append_rows(
            new,
            value_input_option="USER_ENTERED"
        )
        print("Añadidas", len(new), "filas")


# =========================================================
# MAIN
# =========================================================

def run_postprocess():
    script_path = os.path.join(
        os.path.dirname(__file__),
        POSTPROCESS_SCRIPT
    )

    if os.path.exists(script_path):
        subprocess.run(
            [sys.executable, script_path],
            check=True
        )


def main():
    gc = get_gspread_client()
    spreadsheet = open_spreadsheet(gc)

    ws = get_or_create_ws(
        spreadsheet,
        SHEET_YT,
        200000,
        10
    )

    ensure_headers(ws, HEADERS_YT)

    url = read_full1_a1(spreadsheet)

    if not url:
        url = "https://charts.youtube.com/charts/TopSongs/ar/weekly/20180503"

    driver = build_driver()
    extracted_any = False

    try:
        while True:
            print("Procesando", url)

            driver.get(url)
            time.sleep(5)

            accept_cookies_if_present(driver)
            time.sleep(5)

            html = driver.page_source
            date = format_date_from_url(url)
            rows = extract_rows(html, date)

            if not rows:
                print("No se han encontrado filas. Fin del proceso.")
                break

            append_if_new(ws, rows)

            extracted_any = True
            url = next_week_url(url)
            write_full1_a1(spreadsheet, url)

    finally:
        driver.quit()

    if extracted_any:
        run_postprocess()


if __name__ == "__main__":
    main()