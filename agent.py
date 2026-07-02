import sys
import importlib
import importlib.util

import yaml
from pathlib import Path


def main():
    if len(sys.argv) < 2:
        print("Usage: python agent.py <config.yaml>")
        sys.exit(1)

    config_path = sys.argv[1]

    with open(config_path) as f:
        config = yaml.safe_load(f)

    fair_name = config["fair"]["name"]
    platform = config["source"]["platform"]
    exhibitors_url = config["source"]["exhibitors_url"]

    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)

    slug = fair_name.lower().replace(" ", "-")
    scraped_path = str(output_dir / f"{slug}.xlsx")
    enriched_path = str(output_dir / f"{slug}-enriched.xlsx")

    spec = importlib.util.find_spec(f"scrapers.{platform}")
    if spec is None:
        print(f"Error: no scraper for platform '{platform}' (fair: '{fair_name}')")
        sys.exit(1)

    module = importlib.import_module(f"scrapers.{platform}")
    module.scrape(exhibitors_url, scraped_path)

    from enricher import enrich
    enrich(scraped_path, enriched_path)

    print(f"\nDone! Results:")
    print(f"  Scraped:  {scraped_path}")
    print(f"  Enriched: {enriched_path}")


if __name__ == "__main__":
    main()
