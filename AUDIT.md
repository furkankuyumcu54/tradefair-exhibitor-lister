# Audit Report — tradefair-exhibitor-lister

**Date:** 2026-07-02  
**Audited files:** `agent.py`, `scrapers/intermob.py`, `enricher.py`, `fairs/intermob.yaml`  
**Prototype reference:** `/home/qeyosa/fuar_katilimci_listesi/`

---

## 1. Functional Correctness

### 1.1 Does `python agent.py fairs/intermob.yaml` execute the intended workflow?

**Yes.** The workflow is:

1. Read YAML → extract `fair.name` and `source.exhibitors_url`
2. Create `output/` directory
3. Build slug-based filenames (`intermob-2026.xlsx`, `intermob-2026-enriched.xlsx`)
4. Dispatch to `scrapers.intermob.scrape()` via `if/elif`
5. Call `enricher.enrich()` on the scraped output
6. Print result paths

All steps are present and correctly sequenced.

### 1.2 Does `agent.py` only orchestrate?

**Yes.** It reads config, dispatches, and calls enrich. It contains no extraction or enrichment logic. The fair-key dispatch (`intermob`) is hardcoded, but this is an architectural choice (if/elif dispatch) not a functional defect.

### 1.3 Does the scraper still perform the same extraction flow as the original prototype?

**No — and this is by design.** The prototype was a generic multi-agent framework using `requests` + `BeautifulSoup` with strategy pattern. The current scraper is fair-specific using Playwright with hardcoded selectors.

| Aspect | Prototype | Current |
|---|---|---|
| HTTP client | `requests` + BeautifulSoup | Playwright Chromium |
| Listing extraction | Generic HTML strategies | `a.brand-link` → `h2.brand-name`, `p.brand-country` |
| Detail extraction | BS4 meta/links/text | Playwright JS evaluation + regex on contact widget |
| Pagination | `pagination_param` from analysis | `a[href*='page=']` → extract max page number |

The *conceptual* flow (list → paginate → detail → write) is preserved but the implementation is entirely different. This is acceptable given the requirement change from "generic framework" to "fair-specific scraper."

### 1.4 Does the enricher preserve existing values?

**Yes.** The enrichment logic at `enricher.py:433-444` checks:

```python
# Email: only set if row had no email
if email and is_valid_email(email):
    if not row.get("Email") ...
        row["Email"] = email

# Phone: only set if row had no phone
if phone and phone != existing_phone:
    if not existing_phone:
        row["Phone"] = phone

# Website: only set if row had no website
if found_website and not existing_website and found_website_url:
    row["Website"] = found_website_url
```

Existing values are never overwritten.

---

## 2. Prototype Compatibility — Differences Found

Because the prototype was a *generic multi-fair framework* and the current implementation is a *fair-specific tool*, many differences are expected. Below are the specific changes:

### 2.1 Selectors
- **Prototype:** None (generic BS4 extraction from tables/lists/grids)
- **Current:** Hardcoded intermob-specific CSS selectors (`a.brand-link`, `h2.brand-name`, `p.brand-country`, `h4.widget-title`, `div.social a`, `fa-phone`/`fa-globe`/`fa-location-dot` icon markers)
- **Verdict:** New, fair-specific code. Compatible by design intention.

### 2.2 Regexes

| Pattern | Prototype | Current Scraper | Current Enricher |
|---|---|---|---|
| Email | `[\w.][\w.+-]*@[\w-]+(?:\.[\w-]+)+` | *(none)* | `[\w.+-]+@[\w-]+\.[\w.-]+` |
| Phone | `(?:\+?\d{1,3}[-\s.]?)?\(?\d{2,4}\)?...` | `fa-phone">[^<]*</i>\s*([^<]+)` (contact widget) | 2 patterns (expanded from prototype) |
| Booth/Stand | `(?:stand\|booth\|stall)...` | `Stant\s*:?\s*([\dA-Za-z]+)` | N/A |
| Hall | `(?:hall\|salon\|halle\|salon)...` | `Salon\s*:?\s*(\d+)` | N/A |
| Parent company | N/A | `Bağlı Olduğu Firma\s*([^\n]+)` | N/A |

- **Verdict:** Different regexes per design. The enricher's email regex loses the dot-allowed-in-local-part from the prototype (`[\w.][\w.+-]*` → `[\w.+-]+`), dropping support for dots at the start of local parts.

### 2.3 Pagination Logic
- **Prototype:** Uses `pagination_param` and `total_pages_estimate` from the planner's analysis output
- **Current:** Parses `a[href*='page=']` links to find max page, then iterates `?page={n}`
- **Difference:** The current approach discovers pagination at runtime rather than relying on a pre-generated plan. First page uses the URL as-is; subsequent pages append `?page=N`.

### 2.4 Retry Logic
- **Prototype:** No explicit retry in the scraper. The detail scraper catches exceptions silently and returns the original exhibitor.
- **Current:** `retry_page()` with 3 retries, 2-second delay, 30-second timeout per attempt.
- **Difference:** Retry was added. The current approach is more robust.

### 2.5 Checkpoint Behavior
- **Prototype:** No checkpoint in the scraper pipeline.
- **Current:** Text-file checkpoint in `/tmp/` tracking completed slugs. Excel checkpoint save every 50 companies.
- **Difference:** Checkpointing was added. This is an improvement for resume capability.

### 2.6 Extraction Order
- **Prototype:** Listing extraction → Detail enrichment (optional per detail_url)
- **Current:** Listing extraction (Playwright) → Detail extraction (Playwright, same session)
- **Difference:** The current scraper visits detail pages immediately after listing pages, using the same Playwright browser session. The prototype used separate HTTP requests.

### 2.7 Enrichment Workflow
- **Prototype:** Three separate agents (WebsiteDiscoveryAgent → EmailDiscoveryAgent → PhoneDiscoveryAgent), each independent
- **Current:** Single `enrich_company()` function that does website search, then email+phone extraction from homepage + contact pages
- **Difference:** Merged into one pass. The `enrich_company` function returns `(email, phone, website, found_website_bool)` in a single call.

### 2.8 Email Extraction in Scraper
- **Prototype:** Detail page scraper extracts email from `mailto:` links and email regex on text
- **Current scraper:** `extract_detail()` at `scrapers/intermob.py:54-68` initializes `email` to `""` and **never sets it**. No email extraction from detail pages.
- **Impact:** Email is always empty after scraping. The enricher is solely responsible for email. This is functional but differs from the prototype's flow.

---

## 3. Configuration Audit

### 3.1 Values Now Derived from Config/Arguments
| Value | Source |
|---|---|
| Fair name | `fairs/intermob.yaml` → `config["fair"]["name"]` |
| Exhibitors URL | `fairs/intermob.yaml` → `config["source"]["exhibitors_url"]` |
| Scraped output path | Generated from fair name slug + `"-enriched.xlsx"` |
| Enriched output path | Generated from fair name slug + `"-enriched.xlsx"` |

### 3.2 Remaining Hardcoded Values

**`agent.py`:**
| Line | Value | Notes |
|---|---|---|
| 19 | `output` | Output directory path is hardcoded |
| 22 | `lower().replace(" ", "-")` | Slug generation logic |
| 26 | `fair_name.lower().split()[0]` | Fair-key derivation (assumes first word = key) |
| 27 | `"intermob"` | Fair key string |
| 28 | `"scrapers.intermob"` | Module path |

**`scrapers/intermob.py`:**
| Line | Value | Notes |
|---|---|---|
| 10-13 | `FAILED_PAGES`, `TOTAL_PAGES_VISITED`, `CHECKPOINT_SLUGS`, `CHECKPOINT_INTERVAL=50` | Global state and checkpoint interval |
| 34 | `"a.brand-link"` | Listing card selector |
| 34, 42-43 | `"h2.brand-name"`, `"p.brand-country"` | Sub-element selectors |
| 40 | `slug = href.replace("brand/", "")` | Slug extraction (format-specific) |
| 86 | `"h4.widget-title"` | Contact widget selector |
| 88 | `"İletişim"` | Turkish contact widget title |
| 98-108 | 3 regex patterns for phone/address/website | Format-specific |
| 112 | `"div.social a"` | Social links selector |
| 133-143 | 3 regex patterns for salon, stant, parent_company | Format-specific |
| 189 | `max_retries=3` | Retry count |
| 190 | `timeout=30000` | Playwright navigation timeout |
| 194, 257, 300 | `time.sleep(2)`, `time.sleep(1)`, `time.sleep(1)` | Hardcoded delays |
| 217-223 | Viewport, user-agent | Hardcoded |
| 233 | `"a[href*='page=']"` | Pagination link selector |
| 239 | `href.split("page=")[-1].split("&")[0]` | Page number extraction |
| 248 | `f"{exhibitors_url}?page={pnum}"` | Pagination URL construction |
| 305 | `CHECKPOINT_INTERVAL=50` | Excel checkpoint interval |
| 209-210 | `/tmp/{slug_base}_temp.xlsx`, `/tmp/{slug_base}_checkpoint.txt` | Temp file paths |

**`enricher.py`:**
| Line | Value | Notes |
|---|---|---|
| 17-19 | User-agent, Accept, Accept-Language headers | Hardcoded |
| 21 | `REQUEST_TIMEOUT = 15` | HTTP timeout |
| 24 | `PREFERRED_PREFIXES` | 9 email prefixes |
| 27-34 | `CONTACT_PATHS` | 10 contact page paths |
| 164-168 | `search_company_website` queries (3) | Search query templates |
| 167-169 | `"https://lite.duckduckgo.com/lite/"` | Search endpoint |
| 191 | `CONTACT_PATHS` (same as above) | Duplicated in `find_contact_page_urls` |
| 201 | Contact keywords list | 9 keywords |
| 300-301 | `SHORT_WORDS` | 40 2-letter words |
| 318 | TLD list for domain guessing | `[".com.tr", ".com", ...]` |
| 320 | `timeout=5` | Domain guess timeout |
| 454 | `time.sleep(0.8)` | Rate limiting delay |
| 450 | Checkpoint interval every 10 companies | Hardcoded |

---

## 4. Architecture Compliance

| Constraint | Status | Evidence |
|---|---|---|
| No plugin system | ✓ Compliant | No plugin loading or registration |
| No factory pattern | ✓ Compliant | No `StrategyFactory`, `AgentFactory`, etc. |
| No dependency injection | ✓ Compliant | No constructor injection; functions are imported directly |
| No runtime HTML analysis | ✓ Compliant | Selectors are hardcoded strings |
| No runtime selector discovery | ✓ Compliant | No LLM-based or heuristic selector discovery |
| No unnecessary abstraction | ✓ Compliant | No base classes, interfaces, or abstract methods |
| Simple if/elif dispatch | ✓ Compliant | `agent.py:27-29` uses `if fair_key == "intermob":` |
| YAML only contains fair name + exhibitors_url | ✓ Compliant | `fairs/intermob.yaml` has exactly 2 data fields |

**No violations found.**

---

## 5. Project Scope

### 5.1 Scraper Output Columns

Written by `write_excel()` at `scrapers/intermob.py:153-161`:
- Company Name ✓
- Website ✓
- Phone ✓
- Email ✓
- Hall ✓
- Stand ✓
- Detail Page URL ✓

**No scope creep in output**.

### 5.2 Scraper Internal Fields (Not Written)

`extract_detail()` collects these fields that are **never written to Excel**:

| Field | Type | Written? |
|---|---|---|
| `address` | str | No |
| `facebook` | str | No |
| `instagram` | str | No |
| `linkedin` | str | No |
| `twitter` | str | No |
| `country` | str | No |
| `parent_company` | str | No |

These fields are stored in `all_data` via `append` but only the 7 output columns are selected in `write_excel`. This is dead data in memory. **Not scope creep** (not output), but wasteful.

### 5.3 Enricher Modifications

The enricher only writes to:
- Website (only if missing)
- Phone (only if missing)
- Email (only if missing)

**No scope creep in enrichment**.

---

## 6. Code Review

### 6.1 Dead Code

| File | Line(s) | Issue |
|---|---|---|
| `scrapers/intermob.py` | 54-68, 102-148 | `extract_detail` collects `address, facebook, instagram, linkedin, twitter, country, parent_company` but `write_excel` never writes these. They occupy memory across all companies. |
| `enricher.py` | 8 | `from concurrent.futures import ThreadPoolExecutor, as_completed` — imported but never used |
| `enricher.py` | 7 | `from urllib.parse import urljoin` — imported but never used in the top-level namespace (only used inside a closure via `urljoin(base, href)` which Python resolves from the lexical scope anyway — actually it IS used on line 203. So not dead.) |

Actually let me recheck `urljoin` usage in enricher.py. Line 7 imports it, line 203 uses it: `full_url = urljoin(base, href)`. So it IS used.

But `ThreadPoolExecutor` and `as_completed` on line 8 are genuinely unused.

### 6.2 Duplicated Code

| Location | Description |
|---|---|
| `enricher.py:27-34` and `enricher.py:191` | `CONTACT_PATHS` is defined as a module constant and also iterated inside `find_contact_page_urls`. The function `find_contact_page_urls` additionally discovers links from the page but iterates the same `CONTACT_PATHS` list. The constant is used in `find_contact_page_urls` but the paths are essentially the same as what the enricher checks later. Minor duplication. |

### 6.3 Unreachable Code

None found.

### 6.4 Unused Imports

| File | Import | Status |
|---|---|---|
| `enricher.py:8` | `ThreadPoolExecutor, as_completed` | **Unused.** Leftover from prototype's concurrent design. |  

### 6.5 Unused Variables

| File | Variable | Status |
|---|---|---|
| `enricher.py:38-45` | `STATS` dict with 6 keys | Used only in `enrich()` function. Not unused, but is a global mutable accumulator — see risk below. |

### 6.6 Leftover Prototype Artifacts

None in the audited files. The prototype remains in `/home/qeyosa/fuar_katilimci_listesi/` as a separate project, with no imports or files shared.

### 6.7 Temporary / Debug Code

None found.

### 6.8 TODO/FIXME Comments

None found.

---

## 7. Risk Assessment

### CRITICAL

| ID | Description | File:Line | Recommendation |
|---|---|---|---|
| **R1** | **"@" filter in `extract_from_soup` filters ALL emails.** The check `"@" in e` on line 220 evaluates to `True` for every valid email address. Combined with `any( ... )`, this causes every extracted email to be skipped via `continue`. The enricher will never return any emails from website pages. | `enricher.py:220` | Remove `"@"` from the filter list. The intent was likely to filter placeholder emails like `@email.com`, but the bare `@` matches every email. Replace with a more specific check (e.g., `e.endswith("@email.com")`). |

### HIGH

| ID | Description | File:Line | Recommendation |
|---|---|---|---|
| **R2** | **Inconsistent return tuple length in `enrich_company`.** The early return on line 343 returns 3 values `("", "", found_website)`, but the main return on line 383 returns 4 values `(email, phone, website, found_website)`. The caller `enrich()` at line 431 always unpacks 4 values: `email, phone, found_website_url, found_website`. When the early return fires, `found_website_url` will receive `""` and `found_website` will receive `found_website` — but then line 443 checks `if found_website and not existing_website and found_website_url`, where `found_website_url` is `""` which is falsy, so it works accidentally. However, this is fragile. | `enricher.py:343, 383` | Make the early return also return 4 values consistently: `return "", "", "", found_website` |
| **R3** | **DuckDuckGo lite endpoint as sole search source.** The `search_company_website` function at `enricher.py:157` uses `https://lite.duckduckgo.com/lite/` with a direct `requests.post`. This is an undocumented, rate-limited endpoint that can change or block without notice. If it fails silently (caught by `except Exception`), the website search returns empty, and the domain guessing fallback may produce incorrect results. | `enricher.py:157-182` | Add logging for search failures. Consider a fallback search provider or accept this as a known limitation. |
| **R4** | **Domain guessing may assign wrong websites.** The domain guessing logic at `enricher.py:317-337` tries `{domain_base}.com.tr`, `.com`, `.de`, etc. and accepts the first that returns HTTP <400. The verification step (checking if company name words appear in response text) is weak — it only checks the first 3 words longer than 3 characters. This can assign a domain to a company that doesn't own it (e.g., a generic parked domain or a different company with a similar name). | `enricher.py:317-337` | Strengthen verification: require more word matches, check page title, or implement a minimum confidence threshold. |

### MEDIUM

| ID | Description | File:Line | Recommendation |
|---|---|---|---|
| **R5** | **Global mutable `STATS` dict.** The `STATS` dict at module level accumulates state across multiple `enrich()` calls within the same process. If `enrich()` is called twice, counters double. | `enricher.py:38-45` | Initialize stats locally inside `enrich()` and return or print them. |
| **R6** | **Hardcoded `/tmp/` paths for checkpoints.** Both the scraper and enricher use `/tmp/{slug}_checkpoint.txt` and `/tmp/{slug}_temp.xlsx`. On a multi-user system or across different fair runs simultaneously, this could cause collisions. Also, `/tmp/` may be cleaned unexpectedly. | `scrapers/intermob.py:209-210`, `enricher.py:390-391` | Use a project-local temp directory (e.g., `output/.tmp/{slug}/`) instead of `/tmp/`. |
| **R7** | **Missing `__init__.py` in `scrapers/`.** The `scrapers/` directory has no `__init__.py`. While Python 3.3+ supports namespace packages without one, this may cause issues with some tooling or packaging. | `scrapers/__init__.py` missing | Add an empty `__init__.py` file. |
| **R8** | **`requests.post` bypasses session headers in `search_company_website`.** Line 166 calls `requests.post(...)` directly instead of using `SESSION.post(...)`. The request still works but misses the configured headers (User-Agent, Accept, Accept-Language). | `enricher.py:166` | Use `SESSION.post()` instead. |
| **R9** | **Unused `ThreadPoolExecutor` import.** Line 8 imports `ThreadPoolExecutor` and `as_completed` but neither is used in the file. Leftover from the prototype's multi-agent design. | `enricher.py:8` | Remove the unused import. |
| **R10** | **Email never extracted by scraper.** `extract_detail()` initializes `email: ""` and never sets it. The enricher is solely responsible for email. If the enricher fails (e.g., R1), the email column will be entirely empty. | `scrapers/intermob.py:57` | Either extract email from the detail page (e.g., `mailto:` links or contact HTML) or document this as intentional. |

### LOW

| ID | Description | File:Line | Recommendation |
|---|---|---|---|
| **R11** | **Inefficient email filter at `enricher.py:227`.** The version-number filter `intl-tel-input@18.1.5` checks if all domain-part segments are digits. This is fragile — not all version patterns match this heuristic. | `enricher.py:227-233` | If this filter is kept, consider also filtering by common version prefixes (`@v`, `@ver`, `@release`). |
| **R12** | **Hardcoded viewport and user-agent.** The Playwright browser viewport (1920×1080) and user-agent are hardcoded. If the target site changes behavior based on viewport or UA, this is a single point of change. | `scrapers/intermob.py:217-223` | Consider deriving from config or documenting as fair-specific requirements. |
| **R13** | **Checkpoint slug uses company name with lowercasing and transliteration in enricher.** The enricher uses company name as the checkpoint key after aggressive normalization (lowercase, Turkish character replacement). Multiple companies could map to the same slug (e.g., "Çelik" and "Celik" both become "celik"). The enricher would skip the second company. | `enricher.py:418-423` | Use a unique identifier (e.g., row index + name hash) instead of a heavily normalized name. |

---

## Summary

| Category | Count |
|---|---|
| Critical issues | 1 |
| High issues | 3 |
| Medium issues | 5 |
| Low issues | 3 |
| Architecture violations | 0 |
| Scope creep | 0 |

**The single most impactful finding is R1**: the `"@"` filter in `extract_from_soup` (`enricher.py:220`) causes every extracted email to be rejected, rendering email enrichment non-functional. All other issues are correctness, robustness, or maintainability concerns.
