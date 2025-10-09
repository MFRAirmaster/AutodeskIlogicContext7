# Autodesk Forum Scraper Guide

Dette er en guide til at bruge PowerShell scriptet til at scrape Autodesk forum for iLogic kode eksempler.

## Krav

- **Python 3.8+** installeret på din computer
- **PowerShell** (kommer med Windows)
- Internet forbindelse

## Hurtig Start

### 1. Installer Dependencies (Første gang kun)

```powershell
.\Run-ForumScraper.ps1 -InstallDeps
```

Dette installerer:
- Playwright (browser automation)
- aiofiles (async file operations)
- Chromium browser

### 2. Kør Scraper

**Standard (10 sider):**
```powershell
.\Run-ForumScraper.ps1
```

**Specifik antal sider:**
```powershell
.\Run-ForumScraper.ps1 -MaxPages 5
```

**Begræns antal posts:**
```powershell
.\Run-ForumScraper.ps1 -MaxPages 20 -MaxPosts 100
```

**Med custom delay (hvis du får rate-limiting):**
```powershell
.\Run-ForumScraper.ps1 -MaxPages 10 -Delay 3.0
```

## Parametre

| Parameter | Type | Standard | Beskrivelse |
|-----------|------|----------|-------------|
| `-MaxPages` | int | 10 | Antal forum sider at scrape (hver side har ~20-30 posts) |
| `-Delay` | double | 2.0 | Ventetid i sekunder mellem requests |
| `-MaxPosts` | int | 0 (unlimited) | Max antal posts med kode at scrape |
| `-OutputDir` | string | docs/examples/advanced-forum-scraped | Output mappe |
| `-InstallDeps` | switch | false | Installer Python dependencies først |

## Eksempler

### Lille scrape (test)
```powershell
# Scrape kun 3 sider for at teste
.\Run-ForumScraper.ps1 -MaxPages 3
```

### Medium scrape
```powershell
# Scrape 15 sider, max 75 posts
.\Run-ForumScraper.ps1 -MaxPages 15 -MaxPosts 75
```

### Stor scrape
```powershell
# Scrape 50 sider, 3 sekunders delay (mere venlig mod serveren)
.\Run-ForumScraper.ps1 -MaxPages 50 -Delay 3.0
```

### Custom output directory
```powershell
.\Run-ForumScraper.ps1 -MaxPages 10 -OutputDir "C:\MyData\iLogic-Examples"
```

## Output

Scriptet genererer følgende:

### 1. Dokumentation Filer
Placering: `docs/examples/advanced-forum-scraped/`

- `advanced-examples-advanced.md` - Avancerede iLogic eksempler
- `api-examples-advanced.md` - API relaterede eksempler
- `basic-examples-advanced.md` - Basis eksempler
- `troubleshooting-examples-advanced.md` - Fejlfinding eksempler

### 2. Kode Snippets
Placering: `docs/examples/code-snippets/`

- `advanced_[post-id]_[index].vb` - Individuelle VB kode filer
- `api_[post-id]_[index].vb`
- `basic_[post-id]_[index].vb`
- `troubleshooting_[post-id]_[index].vb`

Hver kode fil indeholder:
- Titel på forum post
- URL til original post
- Kategori
- Scraped dato
- Selve koden

## Estimation af Tid

- **3 sider** (~60 posts): ~5-10 minutter
- **10 sider** (~200 posts): ~15-30 minutter
- **20 sider** (~400 posts): ~30-60 minutter
- **50 sider** (~1000 posts): ~2-3 timer

*Tiden afhænger af din internet hastighed og server respons tid.*

## Tips & Best Practices

### 1. Start lille
```powershell
# Test først med få sider
.\Run-ForumScraper.ps1 -MaxPages 2
```

### 2. Brug MaxPosts for kontrol
```powershell
# Scrape kun 50 posts selvom du ser på mange sider
.\Run-ForumScraper.ps1 -MaxPages 20 -MaxPosts 50
```

### 3. Vær venlig mod serveren
```powershell
# Brug højere delay ved store scrapes
.\Run-ForumScraper.ps1 -MaxPages 30 -Delay 3.0
```

### 4. Kør om natten for store jobs
```powershell
# Store scrapes kan tage timer
.\Run-ForumScraper.ps1 -MaxPages 100 -Delay 4.0
```

## Fejlfinding

### Problem: "Python not found"
**Løsning:** Installer Python fra https://www.python.org/downloads/

### Problem: "playwright: command not found"
**Løsning:** 
```powershell
.\Run-ForumScraper.ps1 -InstallDeps
```

### Problem: Rate limiting / timeouts
**Løsning:** Øg delay:
```powershell
.\Run-ForumScraper.ps1 -MaxPages 10 -Delay 5.0
```

### Problem: Scraper stopper midtvejs
**Løsning:** Kør igen - den springer over allerede scrapede posts

## Avanceret Brug

### Schedule automatisk scraping
Opret en Windows Task Scheduler task:

```powershell
# Opret task der kører hver uge
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-File C:\Path\To\Run-ForumScraper.ps1 -MaxPages 10"
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At 2am
Register-ScheduledTask -TaskName "iLogic Forum Scraper" -Action $action -Trigger $trigger
```

### Brug resultater med Context7
Efter scraping kan du bruge de genererede markdown filer som context i dine AI prompts for iLogic udvikling.

## Support

Hvis du oplever problemer:
1. Check at Python er korrekt installeret: `python --version`
2. Installer dependencies igen: `.\Run-ForumScraper.ps1 -InstallDeps`
3. Prøv med færre sider først: `.\Run-ForumScraper.ps1 -MaxPages 2`
4. Øg delay hvis du får timeouts: `-Delay 5.0`

## Yderligere Information

Scriptet bruger:
- **Playwright** for browser automation (håndterer JavaScript dynamisk indhold)
- **AsyncIO** for parallel processing
- **Custom MCP server** for forum scraping capabilities

---

*Genereret som del af AutodeskIlogicContext7 projektet*
