# social-media-leeds

Hittar inlägg där någon söker **statister, skådespelare eller casting** på LinkedIn, X/Twitter, Instagram och Facebook, från **den senaste månaden**. Resultatet sparas som CSV och Excel med kolumnerna `Platform | Title | URL | Snippet | Date`.

Skriptet ersätter det gamla Apps Script-arket. Det söker via Google-resultat, inte direkt i plattformarna.

## Kom igång

Kräver Python 3.9 eller senare.

```powershell
pip install -r requirements.txt
Copy-Item env.example .env
```

1. Skapa ett konto på [serper.dev](https://serper.dev) och kopiera din API-nyckel.
2. Öppna `.env` och klistra in nyckeln efter `SERPER_API_KEY=`.
3. Kör:

```powershell
python scraper.py
```

Resultaten hamnar i mappen `output/`, till exempel `output/social_media_inlagg_2026-09-29.xlsx`. Excel-filen kan importeras direkt i Google Kalkylark (Arkiv → Importera).

## Kostnad

Varje körning gör högst 4 plattformar × `PAGES_PER_QUERY` sidor, alltså 12 sökningar med standardinställningen. Serpers gratisnivå ger 2 500 sökningar, vilket räcker till ungefär 200 körningar.

## Anpassa

Allt ändras högst upp i `scraper.py`, under `INSTÄLLNINGAR`:

| Vad | Var |
|---|---|
| Sökord | `KEYWORDS` |
| Plattformar | `PLATFORMS` |
| Filter på relevanta resultat | `MUST_MATCH_ANY` (reguljära uttryck, hela ord) |
| Antal sidor per plattform | `PAGES_PER_QUERY` i `.env` |

Filtret matchar hela ord, så att `statist` inte släpper igenom inlägg om `statistik`.

## Google Custom Search (äldre alternativ)

Om du fortfarande har nyckeln från det gamla Apps Scriptet kan du sätta `SEARCH_BACKEND=google` och fylla i `GOOGLE_API_KEY` och `GOOGLE_CX`. API:et är stängt för nya kunder och slutar fungera den 1 januari 2027.

## Bra att veta

- **"Senaste månaden" gäller Googles indexering, inte inläggets datum.** Ett gammalt inlägg som Google nyligen har indexerat kan komma med, och nya inlägg som inte är indexerade än saknas.
- **Instagram och Facebook indexeras dåligt av Google.** Räkna med färre träffar därifrån än från LinkedIn och X.
- **Fel stoppar inte körningen.** Vid nätverksfel hoppar skriptet till nästa plattform, och med Ctrl+C sparas det som redan hunnit samlas in.
- **`.env` och `output/` finns i `.gitignore`.** API-nyckeln och resultaten laddas alltså aldrig upp till GitHub.
