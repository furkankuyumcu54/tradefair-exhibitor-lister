# Fix Report

## Real Bugs (Fixed)

| ID | Finding | Fix |
|---|---|---|
| **R1** | `"@"` in the email filter list matched every email (all contain `@`), so `extract_from_soup` discarded all extracted emails. Email enrichment was entirely non-functional. | Removed `"@"` from the filter list at `enricher.py:220`. |
| **R2** | Early return in `enrich_company` (`enricher.py:341`) returned 3 values `("", "", found_website)` but caller always unpacks 4. Would crash with `ValueError` on that code path. | Changed to `return "", "", "", found_website` (4 values). |
| **R8** | `requests.post()` used directly in `search_company_website` instead of `SESSION.post()`, bypassing session headers and cookie persistence. | Replaced with `SESSION.post()`; removed redundant `headers=SESSION.headers`. |
| **R9** | `ThreadPoolExecutor` and `as_completed` imported from `concurrent.futures` but never used. | Removed the unused import. |

## False Positives / Intentional Design (Ignored)

| ID | Finding | Rationale |
|---|---|---|
| **R3** | DuckDuckGo Lite as sole search source | Intentional — acceptable limitation for a fair-specific tool. |
| **R4** | Domain guessing may assign wrong websites | Intentional — YAGNI; known limitation accepted by design. |
| **R5** | Global mutable `STATS` dict | Intentional — simple approach fits project scope. |
| **R6** | Hardcoded `/tmp/` paths for checkpoints | Intentional — simple approach; collisions are unlikely for single runs. |
| **R7** | Missing `__init__.py` in `scrapers/` | Intentional — namespace packages work fine in Python 3.3+. |
| **R10** | Email never extracted by scraper | Intentional — email is exclusively the enricher's responsibility. |

## Changes Made

Only `enricher.py` was modified (4 edits, no new files):
1. Removed `"@"` from the email filter blacklist (R1).
2. Fixed return tuple length on early-exit path (R2).
3. Replaced `requests.post()` with `SESSION.post()` (R8).
4. Removed unused `ThreadPoolExecutor` import (R9).
