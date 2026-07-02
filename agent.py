import sys
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
    exhibitors_url = config["source"]["exhibitors_url"]

    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)

    slug = fair_name.lower().replace(" ", "-")
    scraped_path = str(output_dir / f"{slug}.xlsx")
    enriched_path = str(output_dir / f"{slug}-enriched.xlsx")

    fair_key = fair_name.lower().split()[0]
    if fair_key == "intermob":
        from scrapers.intermob import scrape
        scrape(exhibitors_url, scraped_path)
    else:
        print(f"Error: no scraper for fair '{fair_name}'")
        sys.exit(1)

    from enricher import enrich
    enrich(scraped_path, enriched_path)

    print(f"\nDone! Results:")
    print(f"  Scraped:  {scraped_path}")
    print(f"  Enriched: {enriched_path}")


if __name__ == "__main__":
    main()
