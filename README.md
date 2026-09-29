# social-media-leeds

Hittar inlägg där någon söker **statister, skådespelare eller casting** på LinkedIn, X/Twitter, Instagram och Facebook, från **den senaste månaden**. Resultatet sparas som CSV och Excel med kolumnerna `Platform | Title | URL | Snippet | Date | Source`.

Skriptet ersätter det gamla Apps Script-arket. Det söker via sökmotorer, inte direkt i plattformarna.

## Kom igång (gratis, ingen nyckel)

Kräver Python 3.10 eller senare.

```powershell
pip install -r requirements.txt
Copy-Item env.example .env
python scraper.py
```

Standardinställningen använder DuckDuckGo, som är gratis och inte kräver något konto. Resultaten hamnar i mappen `output/`, till exempel `output/social_media_inlagg_2026-09-29.xlsx`. Excel-filen kan importeras direkt i Google Kalkylark (Arkiv → Importera).

## Sökbackends

Välj med `SEARCH_BACKEND` i `.env`. Flera går att kombinera med kommatecken, till exempel `SEARCH_BACKEND=ddg,exa`. Resultaten slås då ihop och dubbletter tas bort. Kolumnen `Source` visar vilken backend som hittade varje rad.

| Backend | Kostnad | Nyckel | Styrka | Svaghet |
|---|---|---|---|---|
| `ddg` (DuckDuckGo) | Gratis | Nej | Bäst på Facebook-grupper och -inlägg | Olika resultat vid varje körning, kan tillfälligt spärra |
| `exa` | Gratis månadskvot, inget betalkort | `EXA_API_KEY` | Riktigt publiceringsdatum i kolumnen `Date` | Semantisk sökning, inte exakta ord |
| `serper` | 2 500 gratis sökningar | `SERPER_API_KEY` | Googles resultat, stabilt | Tar slut, därefter betalt |
| `google` | – | `GOOGLE_API_KEY` + `GOOGLE_CX` | Samma som gamla arket | Stängt för nya kunder, slutar fungera 2027-01-01 |

**Rekommendation:** `ddg,exa` om allt ska vara gratis. De hittar olika saker, så tillsammans täcker de mer än var för sig.

Nycklar: Exa på [dashboard.exa.ai](https://dashboard.exa.ai), Serper på [serper.dev](https://serper.dev).

## Fulltext med crawl4ai (valfritt)

Sökmotorerna ger bara ett kort utdrag. Med `ENRICH_WITH_CRAWL4AI=true` hämtar skriptet hela annonstexten för **LinkedIn-annonser och -inlägg** och lägger den i kolumnen `Snippet`, upp till 1 500 tecken.

```powershell
pip install -r requirements-crawl4ai.txt
crawl4ai-setup
```

`crawl4ai-setup` laddar ner en webbläsare (Chromium) och behöver bara köras en gång.

X, Instagram och Facebook hämtas **inte**: de visar en inloggningsvägg för robotar, och då blir resultatet sämre än sökmotorns utdrag.

## Anpassa

Allt ändras högst upp i `scraper.py`, under `INSTÄLLNINGAR`:

| Vad | Var |
|---|---|
| Sökord | `KEYWORDS` (och `EXA_QUERY` för Exa) |
| Plattformar | `PLATFORMS` |
| Ord som räcker för att behålla ett resultat | `MUST_MATCH_ANY` |
| Ord som bara räcker i svensk text | `WEAK_MATCH` + `SWEDISH_HINT` |
| Ord som kastar ett resultat | `MUST_NOT_MATCH` |
| Länkar som aldrig är annonser | `EXCLUDE_URL` |

Filtret är gjort för svenska annonser:

- **Hela ord.** `statist` släpper inte igenom `statistik`.
- **Internationella ord kräver svensk text.** `casting`, `extras` och `statist` i singular räknas bara om texten också innehåller å/ä/ö eller ett vanligt svenskt ord. Annars kommer engelska X-inlägg om Hollywood med.
- **Gjutning filtreras bort.** `casting` betyder också gjutning i metallindustrin (till exempel Volvo, Norsk Hydro), och de träffarna kastas.
- **Personprofiler filtreras bort.** Länkar till profiler på LinkedIn (`/in/`) är inga annonser.

## Bra att veta

- **Gratis betyder instabilt.** DuckDuckGo ger olika resultat vid varje körning och visar ibland bara "The site owner hides the web page description" för Facebook. Kör hellre en gång om dagen än att förlita sig på en enda körning.
- **"Senaste månaden" betyder olika saker.** För `ddg`, `serper` och `google` gäller när sökmotorn indexerade sidan, inte inläggets datum. Bara `exa` filtrerar på det riktiga publiceringsdatumet.
- **X ger sällan något.** Svenska castingannonser finns i praktiken på Facebook och LinkedIn.
- **Fel stoppar inte körningen.** Vid nätverksfel eller spärr hoppar skriptet till nästa plattform, och med Ctrl+C sparas det som redan hunnit samlas in.
- **`.env` och `output/` finns i `.gitignore`.** API-nycklarna och resultaten laddas alltså aldrig upp till GitHub.
