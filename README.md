## extract.py

extract.py is a scraper and cache manager built around Screener.in company pages. It defines the Screener base URL, request headers, and timeout settings, and it uses IST-based session bucketing so market metrics are refreshed only when they can actually change. The file includes helpers for Excel sheet parsing, numeric cleaning, and sheet-name sanitization, then loads run workbooks into a structured model of `companies` and `groups`. It also contains page-parsing routines that fetch a company’s Screener profile, extract its displayed name, industry metadata, and top-ratio metrics, and cache those results locally in JSON files with stamped validity rules. The cache logic is careful: company names never expire, industry labels refresh monthly, and price/PE/market-cap metrics are valid only within the current IST market session.

## dashboard.py

dashboard.py is a report generator that reads one or more consolidated Screener workbook files and builds a self-contained interactive HTML dashboard. It discovers files matching `CONSOLIDATED_ALL_SCREENS_*.xlsx`, ignores non-screen or pagination sheets, and parses the remaining sheets to reconstruct company membership, industry data, market cap, P/E, price, and screen-group membership. It enriches the dataset by computing peer benchmarks, earnings yield, discount vs peers, conviction counts, and flags for untrusted P/E values, then it compares runs to identify entrants, exits, movers, and stale snapshots. Finally, it renders a complete `dashboard.html` with embedded CSS, Chart.js-powered charts, filtering controls, sortable tables, change-tracking views, and per-company profile pages.

## how they relate

The two files operate in the same domain but serve different roles. extract.py is focused on extracting raw data from Screener and caching the scraped results, while dashboard.py is focused on ingesting consolidated Excel screen outputs and turning them into an analyst-facing dashboard. Both share the same investment intuition around trusted P/E ranges, peer benchmarking, and company enrichment, but one is about data collection and caching, and the other is about analysis and presentation.
