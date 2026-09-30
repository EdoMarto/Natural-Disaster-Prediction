# Natural Disaster Prediction – Italy

Master's thesis project: a data pipeline that collects Italian open data on natural hazards
(earthquakes, floods, landslides, wildfires, extreme weather, air quality) together with news
articles about weather events, stores everything in a SQL Server database, and classifies the
articles by severity with NLP models.

![Word cloud of the collected news articles](Sentiment/wordcloud.png)

## Architecture

```
 Open data portals ──┐
 (ISPRA, ISTAT,      │   C# / .NET 6 scrapers        SQL Server              Python NLP
 INGV, Protezione    ├──► Selenium, HtmlAgilityPack ──► staging tables ──────► rule-based lexicon
 Civile, MIBACT, …)  │   EPPlus / NPOI (Excel), CSV    (ClimateGeckoDB)       LSTM classifier (Keras)
 News (ANSA) ────────┘                                                         word cloud
```

### 1. Data ingestion (`Scrapers/`, C#)
26 independent scrapers, one per data source. Each one downloads or reads its dataset, normalises it
and inserts it into a dedicated staging table, skipping records that already exist.

| Area | Sources | Scrapers |
|---|---|---|
| Earthquakes | INGV event catalogue, seismic acceleration maps, Civil Protection seismic risk | `File5`, `File7`–`File9` |
| Floods & landslides | ISPRA flood-prone areas, landslide events and hazard, rainfall events | `File10`–`File15`, `File20` |
| Weather extremes | ISTAT precipitation and temperature series, RAN weather stations | `File17`, `File18`, `File22`, `File23`, `File26` |
| Wildfires | Corpo Forestale fire statistics | `File3`, `File4` |
| Territory & environment | ISTAT, Eurostat, PM10, MIBACT, Ministry of Environment, Agenzia per la Coesione Territoriale | `File1`, `File6`, `File16`, `File19`, `File21`, `File24`, `File25` |
| News | ANSA articles searched by keyword (e.g. *maltempo*, *frana*, *terremoto*, *alluvione*) | `File2` |

Dynamic pages are scraped with Selenium and ChromeDriver, static pages with HtmlAgilityPack/ScrapySharp,
and Excel/CSV datasets with EPPlus, NPOI and CsvHelper. All queries are parameterised.

### 2. Article classification (`Sentiment/`, Python)
- `rule_based.py`: scores each headline against a hand-built Italian lexicon of hazard terms (`lexicon.csv`).
- `update_scores.py`: writes the computed scores back to the database.
- `neural_network.py`: trains an LSTM classifier (Embedding → SpatialDropout → LSTM → Dense) that labels
  headlines as *severe* or *not severe*, with early stopping on validation loss.
- `draw_word_cloud.py`: generates the word cloud above from the article corpus.

## Tech stack
**C# / .NET 6** · Selenium · HtmlAgilityPack · EPPlus · NPOI · CsvHelper · **SQL Server** ·
**Python** · TensorFlow/Keras · scikit-learn · NLTK · pandas

## Running it

Requirements: .NET 6 SDK, SQL Server (local or Azure SQL), Google Chrome with a matching
[ChromeDriver](https://chromedriver.chromium.org), Python 3.8–3.10.

Configuration is read from environment variables:

| Variable | Used by | Default |
|---|---|---|
| `CLIMATEGECKO_DB_CONNECTION` | C# scrapers (ADO.NET connection string) and Python scripts (ODBC connection string) | local SQL Server `ClimateGeckoDB` with Windows authentication |
| `CHROMEDRIVER_PATH` | Selenium scrapers | `Settings/` |

```bash
# Scrapers: pick the ones to run in Program.cs
dotnet run

# NLP
cd Sentiment
pip install -r requirements.txt
python rule_based.py
python neural_network.py
```

Input datasets that cannot be downloaded automatically are stored in `Excels/Manuali/`.
`indici_db.docx` documents the database indexes.

## License
[MIT](LICENSE)
