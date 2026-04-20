#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import os
import re
import sys
import time
import subprocess
from datetime import datetime, timedelta

import gspread
from bs4 import BeautifulSoup
from google.oauth2.service_account import Credentials

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

SERVICE_ACCOUNT_FILE = "service_account.json"
SPREADSHEET_ID = "1FoL3lZNvIaC49wC6O_TKp3PNk3fS7YI3FSnZkf1U1iQ"

SHEET_FULL = "Full 1"
SHEET_YT = "YT Argentina"
POSTPROCESS_SCRIPT = "yt_argentina_postprocess_completo.py"

LLISTA_CONST = "YTCHArg"
PAIS_CONST = "Argentina"

HEADERS_YT = ["Núm. Lista", "Cançó", "Interpret", "Data", "Llista", "Pais"]

ABBREVIATIONS = {"dj", "usa", "uk", "vol", "pt", "feat", "ft", "ep"}
LOWERCASE_WORDS = {"x", "&", "feat.", "ft."}


def smart_title_case(text: str) -> str:
    words = text.split()
    out = []
    for w in words:
        lw = w.lower()
        if lw in LOWERCASE_WORDS:
            out.append(lw)
        elif lw.replace(".", "") in ABBREVIATIONS:
            out.append(lw.upper())
        else:
            out.append(lw.capitalize())
    return " ".join(out)


def format_song_title(text: str) -> str:
    words = text.lower().split()
    if words:
        words[0] = words[0].capitalize()
    result = " ".join(words)
    if result.startswith("+") or result.startswith("-"):
        result = "'" + result
    return result


def get_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
    return gspread.authorize(creds)


def get_ws(spreadsheet, name):
    try:
        return spreadsheet.worksheet(name)
    except Exception:
        return spreadsheet.add_worksheet(name, 100000, 10)


def build_driver():
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,3000")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--lang=ca")

    chrome_binary = os.getenv("CHROME_BIN")
    if chrome_binary:
        options.binary_location = chrome_binary

    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)


def extract_date(url: str) -> str:
    m = re.search(r"/weekly/(\d{8})", url)
    if not m:
        raise ValueError(f"No se pudo extraer la fecha de la URL: {url}")
    d = datetime.strptime(m.group(1), "%Y%m%d")
    return d.strftime("%d/%m/%Y")


def next_week(url: str) -> str:
    m = re.search(r"/weekly/(\d{8})", url)
    if not m:
        raise ValueError(f"No se pudo calcular la siguiente semana desde la URL: {url}")
    d = datetime.strptime(m.group(1), "%Y%m%d")
    new = d + timedelta(days=7)
    return url.replace(m.group(1), new.strftime("%Y%m%d"))


def accept_dialogs_if_present(driver):
    selectors = [
        "button[aria-label*='Accept' i]",
        "button[aria-label*='Aceptar' i]",
        "button[aria-label*='Acepta' i]",
        "button[aria-label*='D'acord' i]",
        "button[aria-label*='I agree' i]",
        "button[aria-label*='Got it' i]",
        "tp-yt-paper-button[aria-label*='Accept' i]",
    ]
    for sel in selectors:
        try:
            for btn in driver.find_elements(By.CSS_SELECTOR, sel):
                if btn.is_displayed():
                    driver.execute_script("arguments[0].click();", btn)
                    time.sleep(0.5)
                    return
        except Exception:
            pass


def page_has_chart_content(page_source: str) -> bool:
    html = page_source or ""
    checks = [
        "ytmc-entry-row",
        'id="entity-title"',
        'id="rank"',
        "Visualitzacions setmanals",
        "Setmanes a la llista",
        "Classificació",
    ]
    return any(x in html for x in checks)


def wait_for_rows(driver, timeout: int = 35):
    end = time.time() + timeout
    while time.time() < end:
        accept_dialogs_if_present(driver)
        try:
            rows = driver.find_elements(By.CSS_SELECTOR, "ytmc-entry-row")
            if rows:
                return
            ranks = driver.find_elements(By.CSS_SELECTOR, "#rank")
            titles = driver.find_elements(By.CSS_SELECTOR, "#entity-title")
            if ranks and titles:
                return
        except Exception:
            pass

        if page_has_chart_content(driver.page_source):
            return

        time.sleep(1)

    raise TimeoutException("No se cargaron las filas del chart dentro del tiempo de espera.")


def extract_rows_from_html(page_source: str, weekly_date: str):
    soup = BeautifulSoup(page_source, "html.parser")
    rows_data = []

    for row in soup.select("ytmc-entry-row"):
        rank_el = row.select_one("#rank")
        title_el = row.select_one("#entity-title")
        artist_els = row.select(".artistName")

        rank = rank_el.get_text(" ", strip=True) if rank_el else ""
        title = title_el.get_text(" ", strip=True) if title_el else ""
        artists = " & ".join(
            smart_title_case(a.get_text(" ", strip=True))
            for a in artist_els
            if a.get_text(" ", strip=True)
        )

        if not rank or not title:
            continue

        rows_data.append([
            rank,
            format_song_title(title),
            artists,
            weekly_date,
            LLISTA_CONST,
            PAIS_CONST,
        ])

        if len(rows_data) >= 100:
            break

    return rows_data




def dedupe_best_rank_same_date(rows_data):
    best_by_song_date = {}
    order = []

    for row in rows_data:
        if len(row) < 6:
            continue

        rank_str, song, artist, weekly_date = row[0], row[1], row[2], row[3]
        try:
            rank = int(str(rank_str).strip())
        except Exception:
            continue

        key = (song.strip().casefold(), artist.strip().casefold(), weekly_date.strip())

        if key not in best_by_song_date:
            best_by_song_date[key] = row
            order.append(key)
            continue

        prev_row = best_by_song_date[key]
        try:
            prev_rank = int(str(prev_row[0]).strip())
        except Exception:
            prev_rank = 10**9

        if rank < prev_rank:
            best_by_song_date[key] = row

    return [best_by_song_date[k] for k in order]


def main():
    gc = get_client()
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)
    ws = get_ws(spreadsheet, SHEET_YT)

    if ws.row_values(1) != HEADERS_YT:
        ws.clear()
        ws.update(range_name="A1", values=[HEADERS_YT])

    url = spreadsheet.worksheet(SHEET_FULL).acell("A1").value
    if not url:
        url = "https://charts.youtube.com/charts/TopSongs/do/weekly/20190321"

    driver = build_driver()
    any_rows_added = False

    try:
        while True:
            print("Procesando", url)
            driver.get(url)

            try:
                wait_for_rows(driver)
            except TimeoutException:
                print(f"No hay chart publicado o no se pudo cargar: {url}")
                if not any_rows_added:
                    print("No se añadió ninguna semana nueva. Se mantiene la hoja tal como estaba.")
                break

            time.sleep(2)
            weekly_date = extract_date(url)
            rows = extract_rows_from_html(driver.page_source, weekly_date)
            rows = dedupe_best_rank_same_date(rows)

            if not rows:
                print(f"No se encontraron filas en: {url}")
                if not any_rows_added:
                    print("No se añadió ninguna semana nueva. Se mantiene la hoja tal como estaba.")
                break

            ws.append_rows(rows)
            any_rows_added = True
            url = next_week(url)
            spreadsheet.worksheet(SHEET_FULL).update(range_name="A1", values=[[url]])
    finally:
        driver.quit()

    if any_rows_added:
        subprocess.run([sys.executable, POSTPROCESS_SCRIPT], check=False)
    else:
        print("Postproceso omitido porque no hubo filas nuevas.")


if __name__ == "__main__":
    main()
