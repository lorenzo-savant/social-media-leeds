"""
Cerca post di lavoro per "statist" / "skådespelare" su LinkedIn, X/Twitter,
Instagram e Facebook, SOLO dell'ultimo mese, ed esporta i risultati in
CSV + Excel con le stesse colonne del vecchio foglio Google:
Platform | Title | URL | Snippet

Backend supportati (scegli con SEARCH_BACKEND nel file .env):
  - serper  -> https://serper.dev (risultati Google, 2.500 query gratis)
  - google  -> Google Custom Search JSON API (solo se hai già una chiave:
               chiusa ai nuovi clienti, smette di funzionare il 1/1/2027)
"""

import csv
import os
import time
from datetime import datetime

import requests
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Font

load_dotenv()

# ---------------------------------------------------------------- CONFIG
BACKEND = os.getenv("SEARCH_BACKEND", "serper").lower()
PAGES_PER_QUERY = int(os.getenv("PAGES_PER_QUERY", "3"))  # 10 risultati/pagina

KEYWORDS = [
    "statist",
    "statister",
    "skådespelare",
    "casting",
    "statistjobb",
]

PLATFORMS = {
    "LinkedIn": "site:linkedin.com",
    "Twitter/X": "(site:twitter.com OR site:x.com)",
    "Instagram": "site:instagram.com",
    "Facebook": "site:facebook.com",
}

# Parole che il risultato DEVE contenere (titolo o snippet) per essere tenuto.
# Lascia la lista vuota per non filtrare.
MUST_CONTAIN_ANY = ["statist", "skådespelar", "casting", "söker", "sökes", "extra"]
# ------------------------------------------------------------------------


def search_serper(query: str, page: int) -> list[dict]:
    resp = requests.post(
        "https://google.serper.dev/search",
        headers={"X-API-KEY": os.environ["SERPER_API_KEY"],
                 "Content-Type": "application/json"},
        json={
            "q": query,
            "gl": "se",          # risultati dalla Svezia
            "hl": "sv",          # lingua svedese
            "tbs": "qdr:m",      # <-- SOLO ULTIMO MESE
            "num": 10,
            "page": page,
        },
        timeout=30,
    )
    resp.raise_for_status()
    return [
        {"title": r.get("title", ""), "url": r.get("link", ""),
         "snippet": r.get("snippet", ""), "date": r.get("date", "")}
        for r in resp.json().get("organic", [])
    ]


def search_google_cse(query: str, page: int) -> list[dict]:
    resp = requests.get(
        "https://www.googleapis.com/customsearch/v1",
        params={
            "key": os.environ["GOOGLE_API_KEY"],
            "cx": os.environ["GOOGLE_CX"],
            "q": query,
            "gl": "se",
            "lr": "lang_sv",
            "dateRestrict": "m1",   # <-- SOLO ULTIMO MESE
            "num": 10,
            "start": (page - 1) * 10 + 1,
        },
        timeout=30,
    )
    resp.raise_for_status()
    return [
        {"title": r.get("title", ""), "url": r.get("link", ""),
         "snippet": r.get("snippet", "").replace("\n", " "), "date": ""}
        for r in resp.json().get("items", [])
    ]


SEARCH = {"serper": search_serper, "google": search_google_cse}[BACKEND]


def is_relevant(item: dict) -> bool:
    if not MUST_CONTAIN_ANY:
        return True
    text = f"{item['title']} {item['snippet']}".lower()
    return any(w in text for w in MUST_CONTAIN_ANY)


def main() -> None:
    keyword_block = " OR ".join(f'"{k}"' for k in KEYWORDS)
    seen_urls: set[str] = set()
    rows: list[dict] = []

    for platform, site_filter in PLATFORMS.items():
        query = f"{site_filter} ({keyword_block})"
        print(f"\n🔎 {platform}: {query}")
        for page in range(1, PAGES_PER_QUERY + 1):
            try:
                results = SEARCH(query, page)
            except requests.HTTPError as e:
                print(f"   ⚠️  errore pagina {page}: {e}")
                break
            if not results:
                break
            new = 0
            for r in results:
                if r["url"] in seen_urls or not is_relevant(r):
                    continue
                seen_urls.add(r["url"])
                rows.append({"Platform": platform, "Title": r["title"],
                             "URL": r["url"], "Snippet": r["snippet"],
                             "Date": r["date"]})
                new += 1
            print(f"   pagina {page}: {len(results)} risultati, {new} nuovi")
            time.sleep(1)  # gentile con l'API

    if not rows:
        print("\nNessun risultato trovato.")
        return

    stamp = datetime.now().strftime("%Y-%m-%d")
    os.makedirs("output", exist_ok=True)
    cols = ["Platform", "Title", "URL", "Snippet", "Date"]

    # CSV (utf-8-sig così Excel legge bene å ä ö)
    csv_path = f"output/social_media_inlagg_{stamp}.csv"
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(rows)

    # Excel (si può importare direttamente in Google Sheets)
    xlsx_path = f"output/social_media_inlagg_{stamp}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Risultati"
    ws.append(cols)
    for c in ws[1]:
        c.font = Font(bold=True)
    for row in rows:
        ws.append([row[c] for c in cols])
        link_cell = ws.cell(row=ws.max_row, column=3)
        link_cell.hyperlink = row["URL"]
        link_cell.style = "Hyperlink"
    for col, width in zip("ABCDE", (12, 60, 60, 100, 14)):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"
    wb.save(xlsx_path)

    print(f"\n✅ {len(rows)} risultati salvati in:\n   {csv_path}\n   {xlsx_path}")


if __name__ == "__main__":
    main()
