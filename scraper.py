"""
Söker efter jobbinlägg om "statist" / "skådespelare" på LinkedIn, X/Twitter,
Instagram och Facebook, ENDAST från den senaste månaden, och exporterar
resultaten till CSV + Excel med samma kolumner som det gamla Google-arket:
Platform | Title | URL | Snippet | Date | Source

Sökbackends (välj en eller flera, kommaseparerade, med SEARCH_BACKEND i .env):
  - ddg     -> DuckDuckGo via biblioteket ddgs. GRATIS, ingen nyckel behövs.
  - exa     -> https://exa.ai (gratis månadskvot, ingen betalkort). Ger det
               riktiga publiceringsdatumet för inlägget.
  - serper  -> https://serper.dev (Google-resultat, 2 500 gratis sökningar)
  - google  -> Google Custom Search JSON API (endast om du redan har en nyckel:
               stängd för nya kunder, slutar fungera 2027-01-01)
  Exempel: SEARCH_BACKEND=ddg,exa  -> kör båda och slår ihop resultaten.

Valfritt: ENRICH_WITH_CRAWL4AI=true hämtar varje hittad sida med crawl4ai och
ersätter det korta utdraget med sidans text (se README för begränsningar).
"""

import csv
import os
import re
import sys
import time
from datetime import datetime, timedelta, timezone

import requests
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Font

load_dotenv()

# Windows-terminaler använder ibland cp1252: tvinga UTF-8 så att å ä ö och
# emojis inte kraschar utskriften (t.ex. vid omdirigering till en loggfil).
for stream in (sys.stdout, sys.stderr):
    stream.reconfigure(encoding="utf-8", errors="replace")

# ---------------------------------------------------------------- INSTÄLLNINGAR
BACKENDS_WANTED = [b.strip().lower()
                   for b in os.getenv("SEARCH_BACKEND", "ddg").split(",")
                   if b.strip()]
PAGES_PER_QUERY = int(os.getenv("PAGES_PER_QUERY", "3"))  # 10 resultat/sida
ENRICH = os.getenv("ENRICH_WITH_CRAWL4AI", "false").strip().lower() == "true"
DAYS_BACK = 30  # "senaste månaden"

KEYWORDS = [
    "statist",
    "statister",
    "skådespelare",
    "casting",
    "statistjobb",
]

# Exa är en semantisk sökmotor: den förstår en mening bättre än en OR-lista.
EXA_QUERY = "Vi söker statister eller skådespelare till film, reklam eller tv (casting)"

PLATFORMS = {
    "LinkedIn": ["linkedin.com"],
    "Twitter/X": ["twitter.com", "x.com"],
    "Instagram": ["instagram.com"],
    "Facebook": ["facebook.com"],
}

# Ett resultat behålls om titeln eller utdraget matchar minst ett STARKT ord,
# eller ett SVAGT ord tillsammans med en svensk ledtråd. "Casting" och "extras"
# är internationella ord: utan kravet kommer engelskspråkiga X-inlägg om vad
# som helst med. Hela ord, så att "statist" INTE träffar "statistik".
# ("söker"/"sökes" räcker inte ensamt: det träffar alla jobbannonser.)
# Lämna båda listorna tomma för att inte filtrera alls.
MUST_MATCH_ANY = [
    r"\bstatist(er|erna|en|jobb\w*)\b",
    r"\bskådespel\w*",
]
WEAK_MATCH = [
    r"\bstatist\b",       # på engelska betyder "statist" något politiskt
    r"\bcasting\w*",
    r"\bextras?\b",
]
SWEDISH_HINT = r"[åäö]|\b(och|söker|sökes|till|för|med|inspelning|ansök\w*)\b"

# ... men kastas om det matchar något av dessa. "Casting" betyder också
# gjutning (metallindustri), och de träffarna är många.
MUST_NOT_MATCH = [
    r"\bgjut\w*",
    r"\b(die|sand|investment|metal|aluminium|aluminum|steel|iron)[ -]casting\b",
    r"\bcasting\s+(engineer|manager|plant|process|operator|technician)\w*",
    r"\b(engineering|technology|department)\s+manager\b",
    r"\bfoundry\b",
]

# Sidor som aldrig är inlägg eller annonser (t.ex. personprofiler på LinkedIn).
EXCLUDE_URL = [
    r"linkedin\.com/in/",
    r"linkedin\.com/company/[^/]+/?$",
]
# ------------------------------------------------------------------------------

RELEVANT_RE = [re.compile(p, re.IGNORECASE) for p in MUST_MATCH_ANY]
WEAK_RE = [re.compile(p, re.IGNORECASE) for p in WEAK_MATCH]
SWEDISH_RE = re.compile(SWEDISH_HINT, re.IGNORECASE)
IRRELEVANT_RE = [re.compile(p, re.IGNORECASE) for p in MUST_NOT_MATCH]
EXCLUDE_URL_RE = [re.compile(p, re.IGNORECASE) for p in EXCLUDE_URL]
KEYWORD_BLOCK = " OR ".join(f'"{k}"' for k in KEYWORDS)


def google_style_query(domains: list[str]) -> str:
    sites = " OR ".join(f"site:{d}" for d in domains)
    if len(domains) > 1:
        sites = f"({sites})"
    return f"{sites} ({KEYWORD_BLOCK})"


# Varje backend tar emot plattformens domäner och returnerar en lista med
# {"title", "url", "snippet", "date"}. Paginering sköts inne i funktionen.

def search_serper(domains: list[str]) -> list[dict]:
    out = []
    for page in range(1, PAGES_PER_QUERY + 1):
        resp = requests.post(
            "https://google.serper.dev/search",
            headers={"X-API-KEY": os.environ["SERPER_API_KEY"],
                     "Content-Type": "application/json"},
            json={
                "q": google_style_query(domains),
                "gl": "se",          # resultat från Sverige
                "hl": "sv",          # svenska
                "tbs": "qdr:m",      # <-- ENDAST SENASTE MÅNADEN
                "num": 10,
                "page": page,
            },
            timeout=30,
        )
        resp.raise_for_status()
        items = resp.json().get("organic", [])
        if not items:
            break
        out += [{"title": r.get("title", ""), "url": r.get("link", ""),
                 "snippet": r.get("snippet", ""), "date": r.get("date", "")}
                for r in items]
        time.sleep(1)  # var snäll mot API:et
    return out


def search_google_cse(domains: list[str]) -> list[dict]:
    out = []
    for page in range(1, PAGES_PER_QUERY + 1):
        resp = requests.get(
            "https://www.googleapis.com/customsearch/v1",
            params={
                "key": os.environ["GOOGLE_API_KEY"],
                "cx": os.environ["GOOGLE_CX"],
                "q": google_style_query(domains),
                "gl": "se",
                "lr": "lang_sv",
                "dateRestrict": "m1",   # <-- ENDAST SENASTE MÅNADEN
                "num": 10,
                "start": (page - 1) * 10 + 1,
            },
            timeout=30,
        )
        resp.raise_for_status()
        items = resp.json().get("items", [])
        if not items:
            break
        out += [{"title": r.get("title", ""), "url": r.get("link", ""),
                 "snippet": r.get("snippet", "").replace("\n", " "), "date": ""}
                for r in items]
        time.sleep(1)
    return out


def search_ddg(domains: list[str]) -> list[dict]:
    from ddgs import DDGS  # importeras här så att andra backends funkar utan ddgs
    from ddgs.exceptions import DDGSException

    out = []
    # DuckDuckGo hanterar inte "site:a OR site:b" pålitligt: en sökning per domän.
    for domain in domains:
        try:
            items = DDGS().text(
                f"site:{domain} ({KEYWORD_BLOCK})",
                region="se-sv",
                timelimit="m",          # <-- ENDAST SENASTE MÅNADEN
                max_results=PAGES_PER_QUERY * 10,
            ) or []
        except DDGSException as e:
            # ddgs kastar ett undantag när det inte finns några träffar.
            if "no results" not in str(e).lower():
                raise
            items = []
        out += [{"title": r.get("title", ""), "url": r.get("href", ""),
                 "snippet": r.get("body", ""), "date": ""}
                for r in items]
        time.sleep(2)  # DuckDuckGo spärrar snabbt vid för många anrop
    return out


def search_exa(domains: list[str]) -> list[dict]:
    since = datetime.now(timezone.utc) - timedelta(days=DAYS_BACK)
    resp = requests.post(
        "https://api.exa.ai/search",
        headers={"x-api-key": os.environ["EXA_API_KEY"],
                 "Content-Type": "application/json"},
        json={
            "query": EXA_QUERY,
            "type": "auto",
            "numResults": PAGES_PER_QUERY * 10,
            "includeDomains": domains,
            # <-- ENDAST SENASTE MÅNADEN, på inläggets publiceringsdatum
            "startPublishedDate": since.strftime("%Y-%m-%dT%H:%M:%SZ"),
            "contents": {"highlights": {"maxCharacters": 400}},
        },
        timeout=60,
    )
    resp.raise_for_status()
    return [
        {"title": r.get("title") or "", "url": r.get("url", ""),
         "snippet": " … ".join(r.get("highlights") or []),
         "date": (r.get("publishedDate") or "")[:10]}
        for r in resp.json().get("results", [])
    ]


BACKENDS = {
    "ddg": (search_ddg, []),
    "exa": (search_exa, ["EXA_API_KEY"]),
    "serper": (search_serper, ["SERPER_API_KEY"]),
    "google": (search_google_cse, ["GOOGLE_API_KEY", "GOOGLE_CX"]),
}


def check_config() -> None:
    """Stoppar direkt med ett begripligt meddelande om .env är fel ifylld."""
    if not BACKENDS_WANTED:
        sys.exit("❌ SEARCH_BACKEND är tom i .env-filen.")
    unknown = [b for b in BACKENDS_WANTED if b not in BACKENDS]
    if unknown:
        sys.exit(f"❌ Okänd SEARCH_BACKEND: {', '.join(unknown)}. "
                 f"Tillåtna värden: {', '.join(BACKENDS)}.")
    missing = [k for b in BACKENDS_WANTED for k in BACKENDS[b][1]
               if not os.getenv(k, "").strip()]
    if missing:
        sys.exit(f"❌ Saknar {', '.join(missing)} i .env-filen "
                 f"(kopiera env.example till .env och fyll i nyckeln).")
    if "ddg" in BACKENDS_WANTED:
        try:
            import ddgs  # noqa: F401
        except ImportError:
            sys.exit("❌ Backend 'ddg' kräver biblioteket ddgs: pip install -r requirements.txt")
    if ENRICH:
        try:
            import crawl4ai  # noqa: F401
        except ImportError:
            sys.exit("❌ ENRICH_WITH_CRAWL4AI=true kräver crawl4ai: "
                     "pip install -r requirements-crawl4ai.txt && crawl4ai-setup")


def is_relevant(item: dict) -> bool:
    if any(p.search(item["url"]) for p in EXCLUDE_URL_RE):
        return False
    text = f"{item['title']} {item['snippet']}"
    if any(p.search(text) for p in IRRELEVANT_RE):
        return False
    if not RELEVANT_RE and not WEAK_RE:
        return True
    if any(p.search(text) for p in RELEVANT_RE):
        return True
    return any(p.search(text) for p in WEAK_RE) and bool(SWEDISH_RE.search(text))


def normalize_url(url: str) -> str:
    """Samma inlägg kan dyka upp med olika varianter av länken."""
    url = url.split("#")[0].split("?")[0].rstrip("/")
    return re.sub(r"^https?://(www\.|m\.|mobile\.|[a-z]{2}\.)?", "", url.lower())


def collect(rows: list[dict]) -> None:
    """Fyller på rows under körningen, så att inget går förlorat vid avbrott."""
    seen_urls: set[str] = set()

    for backend in BACKENDS_WANTED:
        search = BACKENDS[backend][0]
        for platform, domains in PLATFORMS.items():
            print(f"\n🔎 [{backend}] {platform}")
            try:
                results = search(domains)
            except Exception as e:  # nätverksfel, API-fel, spärr från DDG ...
                # Hoppa till nästa plattform men behåll allt som redan samlats in.
                print(f"   ⚠️  fel: {e}")
                continue
            new = 0
            for r in results:
                key = normalize_url(r["url"])
                if not r["url"] or key in seen_urls or not is_relevant(r):
                    continue
                seen_urls.add(key)
                rows.append({"Platform": platform, "Title": r["title"],
                             "URL": r["url"], "Snippet": r["snippet"],
                             "Date": r["date"], "Source": backend})
                new += 1
            print(f"   {len(results)} resultat, {new} nya")


# Sidor som crawl4ai kan läsa utan inloggning, och CSS-väljaren för själva
# texten (annars kommer menyer och länkar med). X, Instagram och Facebook
# visar en inloggningsvägg för robotar, så de hämtas inte.
ENRICH_TARGETS = {
    r"linkedin\.com/jobs/view/": ".show-more-less-html__markup",
    r"linkedin\.com/posts/": ".attributed-text-segment-list__content",
}


def enrich_with_crawl4ai(rows: list[dict]) -> None:
    """Hämtar annonsens fulltext och ersätter det korta utdraget med den."""
    import asyncio
    from crawl4ai import AsyncWebCrawler, BrowserConfig, CacheMode, CrawlerRunConfig

    todo = []
    for row in rows:
        for pattern, selector in ENRICH_TARGETS.items():
            if re.search(pattern, row["URL"]):
                todo.append((row, selector))
                break
    if not todo:
        print("\n🕷️  crawl4ai: inga sidor att hämta (bara LinkedIn-annonser/-inlägg stöds).")
        return

    async def run() -> None:
        async with AsyncWebCrawler(config=BrowserConfig(headless=True,
                                                        verbose=False)) as crawler:
            for i, (row, selector) in enumerate(todo, 1):
                print(f"   🕷️  {i}/{len(todo)} {row['URL'][:80]}")
                cfg = CrawlerRunConfig(cache_mode=CacheMode.BYPASS,
                                       css_selector=selector, page_timeout=30000,
                                       verbose=False)
                try:
                    res = await crawler.arun(url=row["URL"], config=cfg)
                except Exception as e:
                    print(f"      ⚠️  {e}")
                    continue
                text = str(res.markdown or "") if res.success else ""
                text = re.sub(r"\s+", " ", re.sub(r"[#*_>`]", " ", text)).strip()
                # Tom eller kort text (spärr, ändrad sidlayout): behåll utdraget.
                if len(text) > len(row["Snippet"]):
                    row["Snippet"] = text[:1500]
                time.sleep(1)

    print(f"\n🕷️  Hämtar fulltext för {len(todo)} LinkedIn-sidor med crawl4ai ...")
    asyncio.run(run())


def export(rows: list[dict]) -> None:
    stamp = datetime.now().strftime("%Y-%m-%d")
    os.makedirs("output", exist_ok=True)
    cols = ["Platform", "Title", "URL", "Snippet", "Date", "Source"]

    # CSV (utf-8-sig så att Excel visar å ä ö korrekt)
    csv_path = f"output/social_media_inlagg_{stamp}.csv"
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(rows)

    # Excel (kan importeras direkt i Google Kalkylark)
    xlsx_path = f"output/social_media_inlagg_{stamp}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultat"
    ws.append(cols)
    for c in ws[1]:
        c.font = Font(bold=True)
    for row in rows:
        ws.append([row[c] for c in cols])
        link_cell = ws.cell(row=ws.max_row, column=3)
        link_cell.hyperlink = row["URL"]
        link_cell.style = "Hyperlink"
    for col, width in zip("ABCDEF", (12, 60, 60, 100, 12, 9)):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"
    wb.save(xlsx_path)

    print(f"\n✅ {len(rows)} resultat sparade i:\n   {csv_path}\n   {xlsx_path}")


def main() -> None:
    check_config()
    print(f"Sökbackends: {', '.join(BACKENDS_WANTED)}"
          f"{' + crawl4ai' if ENRICH else ''}")
    rows: list[dict] = []
    try:
        collect(rows)
        if ENRICH and rows:
            enrich_with_crawl4ai(rows)
    except KeyboardInterrupt:
        print("\n⏹  Avbruten – sparar det som hunnit samlas in.")
    if not rows:
        print("\nInga resultat hittades.")
        return
    export(rows)


if __name__ == "__main__":
    main()
