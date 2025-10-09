# Autodesk iLogic Context7

Et omfattende system til at samle, organisere og søge i iLogic kode eksempler fra Autodesk Inventor Programming Forum.

## 🚀 Funktioner

### **Parallel Forum Scraping**
- **10 samtidige workers** for maksimal hastighed
- **Intelligent rate limiting** for at undgå blokering
- **Automatisk kategorisering** af kode eksempler
- **SQLite database** til hurtig søgning

### **MCP Server Integration**
- **5 forskellige tools** til forskellige opgaver
- **Lokalt søgning** i scraped data (ingen live scraping nødvendig)
- **Kategoriseret søgning** (basic, advanced, api, troubleshooting)
- **Relaterede eksempler** baseret på kode snippets

### **Omfattende Kodebase**
- **254+ kode eksempler** fra forum posts
- **4 kategorier** organiseret efter kompleksitet
- **Individuelle VB filer** for hver kode snippet
- **Markdown dokumentation** med eksempler
- **Nye eksempler**: Part rule med reference parameters, DXF export metoder (2025 kompatible)

## 📁 Projekt Struktur

```
autodesk-ilogic-context7/
├── autodesk-forum-scraper/          # MCP server
│   ├── index.js                     # Server implementation
│   └── package.json                 # Dependencies
├── docs/examples/                   # Scraped data
│   ├── ilogic-snippets.db          # SQLite database
│   ├── code-snippets/              # Individual VB files
│   ├── parallel-forum-scraped/     # Categorized docs
│   └── *-examples-*.md            # Documentation files
├── parallel-forum-scraper.py       # Main scraper
├── mcp.json                        # MCP configuration
└── README.md                       # This file
```

## 🛠️ Installation & Setup

### 1. Install Dependencies
```bash
# Install Python dependencies
pip install playwright aiofiles

# Install Node.js dependencies for MCP server
cd autodesk-forum-scraper
npm install
cd ..
```

### 2. Konfigurer MCP Server
MCP konfigurationen er allerede sat op i `mcp.json`. Context7 vil automatisk starte MCP serveren.

### 3. Kør Scraping (Valgfrit)
Hvis du vil opdatere kode eksemplerne:
```bash
python parallel-forum-scraper.py
```

## 🔍 MCP Server Tools

### **search_local_snippets**
Søger i lokal database efter kode eksempler.
```json
{
  "query": "parameter",
  "category": "advanced",
  "limit": 10
}
```

### **get_code_examples**
Henter eksempler fra en specifik kategori.
```json
{
  "category": "api",
  "limit": 20
}
```

### **get_related_examples**
Finder relaterede eksempler baseret på kode.
```json
{
  "code_snippet": "Dim oDoc As Document = ThisApplication.ActiveDocument",
  "limit": 5
}
```

### **scrape_forum_post** (Live)
Scraper en individuel forum post.
```json
{
  "url": "https://forums.autodesk.com/t5/...",
  "extractCode": true
}
```

### **scrape_forum_list** (Live)
Scraper liste af forum posts.
```json
{
  "url": "https://forums.autodesk.com/t5/...",
  "maxPosts": 10
}
```

## 📊 Database Schema

### **posts** tabel
- `post_id`: Unik post ID
- `title`: Post titel
- `author`: Forfatter
- `content`: Post indhold
- `category`: Kategori (basic/advanced/api/troubleshooting)
- `url`: Forum URL
- `scraped_at`: Scrape timestamp

### **code_snippets** tabel
- `post_id`: Reference til post
- `snippet_index`: Kode blok index
- `code`: VB kode
- `language`: Sprog (vb/csharp)
- `category`: Nedarvet kategori

## 🎯 Kategorier

### **Basic** (4 eksempler)
Enkle iLogic regler og grundlæggende funktioner.

### **Advanced** (186 eksempler)
Komplekse automation, geometri manipulation, event triggers.

### **API** (56 eksempler)
Inventor API integration, COM objekter, advanced programming.

### **Troubleshooting** (8 eksempler)
Fejlhåndtering, debugging, problemløsning.

## 🚀 Performance Metrics

| Metric | Parallel (5 workers) | Forbedring |
|--------|---------------------|------------|
| Scraping tid | 18 min | 4.5x hurtigere |
| Success rate | 45.7% | Robust rate limiting |
| Kode eksempler | 254+ | Omfattende bibliotek |
| Søge hastighed | < 100ms | SQLite indeks |

## 🔧 Udvikling

### Tilføj Ny Kategori
1. Opdater `categorize_post()` metoden i scraper
2. Kør scraping igen
3. Opdater dokumentation

### Udvid MCP Server
1. Tilføj nye tools i `setupToolHandlers()`
2. Implementer metoder i `AutodeskForumScraperServer` klasse
3. Test med Context7

## 📈 Fremtidige Forbedringer

- [ ] **AI-drevet kategorisering** med machine learning
- [ ] **Code similarity search** for bedre relaterede resultater
- [ ] **Automated testing** af kode eksempler
- [ ] **Integration med Inventor** direkte fra Context7
- [ ] **User contributions** system for community kode

## 🤝 Bidrag

Dette er et open source projekt. Bidrag er velkomne!

1. Fork repository
2. Opret feature branch
3. Implementer ændringer
4. Test grundigt
5. Submit pull request

## 📄 Licens

ISC License - se LICENSE fil for detaljer.

---

**Bygget med ❤️ for Autodesk Inventor community**
