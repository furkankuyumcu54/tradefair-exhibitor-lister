import time
import re
import os
import shutil

from playwright.sync_api import sync_playwright
from openpyxl import Workbook


FAILED_PAGES = []
TOTAL_PAGES_VISITED = 0
CHECKPOINT_SLUGS = set()
CHECKPOINT_INTERVAL = 50


def load_checkpoint(checkpoint_file):
    global CHECKPOINT_SLUGS
    try:
        with open(checkpoint_file, "r", encoding="utf-8") as f:
            for line in f:
                slug = line.strip()
                if slug:
                    CHECKPOINT_SLUGS.add(slug)
    except FileNotFoundError:
        CHECKPOINT_SLUGS = set()


def save_checkpoint(slug, checkpoint_file):
    with open(checkpoint_file, "a", encoding="utf-8") as f:
        f.write(slug + "\n")


def extract_company_from_listing(page):
    cards = page.query_selector_all("a.brand-link")
    companies = []
    for card in cards:
        href = card.get_attribute("href")
        if not href:
            continue
        slug = href.replace("brand/", "")
        name_el = card.query_selector("h2.brand-name")
        country_el = card.query_selector("p.brand-country")
        name = name_el.inner_text().strip() if name_el else ""
        country = country_el.inner_text().strip() if country_el else ""
        companies.append({
            "slug": slug,
            "name": name,
            "country": country,
        })
    return companies


def extract_detail(page, company, detail_url):
    data = {
        "company_name": company["name"],
        "website": "",
        "email": "",
        "phone": "",
        "address": "",
        "facebook": "",
        "instagram": "",
        "linkedin": "",
        "country": company["country"],
        "salon": "",
        "stant": "",
        "parent_company": "",
        "detail_url": detail_url,
    }

    try:
        page.wait_for_selector("h1", timeout=10000)
    except Exception:
        pass

    h1_el = page.query_selector("h1")
    if h1_el:
        h1_text = h1_el.inner_text().strip()
        h1_text = re.sub(r'\s+Temsilci\s+Firma\s*$', '', h1_text).strip()
        if h1_text:
            data["company_name"] = h1_text

    page_text = page.inner_text("body")

    contact_html = page.evaluate("""
    function() {
        var h4s = document.querySelectorAll('h4.widget-title');
        for (var i = 0; i < h4s.length; i++) {
            if (h4s[i].textContent.indexOf('\u0130leti\u015fim') !== -1) {
                var parent = h4s[i].parentElement;
                if (parent) return parent.innerHTML;
            }
        }
        return '';
    }
    """)

    if contact_html:
        phone_match = re.search(r'fa-phone">[^<]*</i>\s*([^<]+)', contact_html)
        if phone_match:
            data["phone"] = phone_match.group(1).strip()

        addr_match = re.search(r'fa-location-dot">[^<]*</i>\s*([^<]+)', contact_html)
        if addr_match:
            data["address"] = addr_match.group(1).strip()

        web_match = re.search(r'fa-globe">[^<]*</i>\s*<a\s+href="([^"]+)"', contact_html)
        if web_match:
            data["website"] = web_match.group(1).strip()

    social_links = page.evaluate("""
    function() {
        var links = document.querySelectorAll('div.social a');
        var results = [];
        for (var i = 0; i < links.length; i++) {
            var href = links[i].getAttribute('href');
            if (href && href !== '#') results.push(href);
        }
        return results;
    }
    """)
    for link in social_links:
        link_lower = link.lower()
        if "facebook.com" in link_lower or "fb.com" in link_lower:
            data["facebook"] = link
        elif "instagram.com" in link_lower or "instagr" in link_lower:
            data["instagram"] = link
        elif "linkedin.com" in link_lower or "linkedin" in link_lower:
            data["linkedin"] = link
        elif "twitter.com" in link_lower or "x.com" in link_lower:
            if not data.get("twitter"):
                data["twitter"] = link

    salon_match = re.search(r'Salon\s*:?\s*(\d+)', page_text)
    if salon_match:
        data["salon"] = salon_match.group(1).strip()

    stant_match = re.search(r'Stant\s*:?\s*([\dA-Za-z]+)', page_text)
    if stant_match:
        data["stant"] = stant_match.group(1).strip()

    parent_match = re.search(r'Ba\u011fl\u0131 Oldu\u011fu Firma\s*([^\n]+)', page_text)
    if parent_match:
        data["parent_company"] = parent_match.group(1).strip()

    return data


def write_excel(all_data, filepath):
    wb = Workbook()
    ws = wb.active
    ws.title = "Exhibitors"

    headers = [
        "Company Name",
        "Website",
        "Phone",
        "Email",
        "Hall",
        "Stand",
        "Detail Page URL",
    ]
    ws.append(headers)

    for row in all_data:
        ws.append([
            row.get("company_name", ""),
            row.get("website", ""),
            row.get("phone", ""),
            row.get("email", ""),
            row.get("salon", ""),
            row.get("stant", ""),
            row.get("detail_url", ""),
        ])

    time.sleep(0.5)
    wb.save(filepath)
    print(f"Excel checkpoint saved to {filepath}")


def copy_to_final(temp_excel, final_excel):
    if os.path.exists(temp_excel):
        time.sleep(0.5)
        shutil.copy2(temp_excel, final_excel)
        print(f"Copied to final location: {final_excel}")


def retry_page(page, url, max_retries=3):
    for attempt in range(max_retries):
        try:
            page.goto(url, wait_until="networkidle", timeout=30000)
            return True
        except Exception as e:
            print(f"  Retry {attempt + 1}/{max_retries} for {url}: {e}")
            time.sleep(2)
    return False


def scrape(exhibitors_url, output_path):
    global TOTAL_PAGES_VISITED, FAILED_PAGES, CHECKPOINT_SLUGS
    FAILED_PAGES = []
    TOTAL_PAGES_VISITED = 0
    CHECKPOINT_SLUGS = set()

    from urllib.parse import urlparse
    parsed = urlparse(exhibitors_url)
    base_url = f"{parsed.scheme}://{parsed.netloc}"

    slug_base = os.path.splitext(os.path.basename(output_path))[0]
    temp_excel = f"/tmp/{slug_base}_temp.xlsx"
    checkpoint_file = f"/tmp/{slug_base}_checkpoint.txt"

    load_checkpoint(checkpoint_file)
    all_data = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            viewport={"width": 1920, "height": 1080},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
        )
        page = context.new_page()

        print("Discovering total page count...")
        if not retry_page(page, exhibitors_url):
            print("Failed to load initial page!")
            browser.close()
            return

        page_links = page.query_selector_all("a[href*='page=']")
        max_page = 1
        for pl in page_links:
            href = pl.get_attribute("href")
            if href and "page=" in href:
                try:
                    pnum = int(href.split("page=")[-1].split("&")[0])
                    max_page = max(max_page, pnum)
                except (ValueError, IndexError):
                    pass
        print(f"Total pages detected: {max_page}")

        seen_slugs = set()
        all_companies = []
        for pnum in range(1, max_page + 1):
            url = exhibitors_url if pnum == 1 else f"{exhibitors_url}?page={pnum}"
            TOTAL_PAGES_VISITED += 1
            print(f"\n--- Listing Page {pnum}/{max_page} ---")

            if not retry_page(page, url):
                FAILED_PAGES.append(url)
                print(f"  FAILED to load {url}")
                continue

            time.sleep(1)
            companies = extract_company_from_listing(page)
            deduped = 0
            for c in companies:
                if c["slug"] not in seen_slugs:
                    seen_slugs.add(c["slug"])
                    all_companies.append(c)
                else:
                    deduped += 1
            print(f"  Found {len(companies)} companies ({deduped} duplicates skipped)")

        print(f"\nTotal unique companies: {len(all_companies)}")

        for idx, company in enumerate(all_companies):
            slug = company["slug"]
            if slug in CHECKPOINT_SLUGS:
                print(f"  [{idx + 1}/{len(all_companies)}] SKIP (checkpoint): {company['name']}")
                continue

            url = f"{base_url}/brand/{slug}"
            print(f"  [{idx + 1}/{len(all_companies)}] Scraping: {company['name']}")

            if not retry_page(page, url):
                FAILED_PAGES.append(url)
                data = {
                    "company_name": company["name"],
                    "website": "",
                    "email": "",
                    "phone": "",
                    "address": "",
                    "facebook": "",
                    "instagram": "",
                    "linkedin": "",
                    "country": company["country"],
                    "salon": "",
                    "stant": "",
                    "parent_company": "",
                    "detail_url": url,
                }
                all_data.append(data)
                save_checkpoint(slug, checkpoint_file)
                continue

            time.sleep(1)
            data = extract_detail(page, company, url)
            all_data.append(data)
            save_checkpoint(slug, checkpoint_file)

            if (idx + 1) % CHECKPOINT_INTERVAL == 0:
                print(f"  *** Checkpoint: saving Excel after {idx + 1} companies ***")
                try:
                    write_excel(all_data, temp_excel)
                except Exception as e:
                    print(f"  Warning: checkpoint save failed: {e}")

        browser.close()

    write_excel(all_data, temp_excel)
    copy_to_final(temp_excel, output_path)

    print("\n" + "=" * 60)
    print(f"Total exhibitors scraped: {len(all_data)}")
    print(f"Total pages visited: {TOTAL_PAGES_VISITED}")
    if FAILED_PAGES:
        print(f"Failed pages ({len(FAILED_PAGES)}):")
        for fp in FAILED_PAGES:
            print(f"  - {fp}")
    else:
        print("Failed pages: None")
    print(f"File location: {output_path}")
    print("=" * 60)
