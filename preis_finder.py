"""Sucht Kaufangebote zu den Begriffen aus produkte.txt und exportiert sie nach Excel.

Die Suche fragt ohne API-Schlüssel die öffentlich sichtbaren Suchergebnisse von
Amazon.de ab. Preise und Verfügbarkeit können sich bis zum Öffnen des Links
ändern; die Shop-Seite ist daher die maßgebliche Quelle.
"""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup


BASE_DIR = Path(__file__).resolve().parent
PRODUCTS_FILE = BASE_DIR / "produkte.txt"
OUTPUT_FILE = BASE_DIR / "ergebnisse" / "angebote.xlsx"
SEARCH_URL = "https://www.amazon.de/s"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/122.0 Safari/537.36"
    ),
    "Accept-Language": "de-DE,de;q=0.9,en;q=0.7",
}
COLUMNS = ["Suchbegriff", "Produkt-Info", "Preis", "Verkäufer", "Link", "Datum"]


def read_products(path: Path) -> list[str]:
    """Liest nichtleere Zeilen; die optionale Markdown-Kennung 'Plaintext' wird ignoriert."""
    if not path.exists():
        raise FileNotFoundError(f"Produktdatei nicht gefunden: {path}")
    return [
        line.strip()
        for line in path.read_text(encoding="utf-8").splitlines()
        if line.strip() and line.strip().lower() != "plaintext"
    ]


def seller_from_url(url: str) -> str:
    host = urlparse(url).netloc.lower().removeprefix("www.")
    return host or "Unbekannter Verkäufer"


def find_offers(search_term: str, limit: int = 10) -> list[dict[str, str]]:
    """Liest die aktuellen, in den Amazon-Suchergebnissen sichtbaren Angebote aus."""
    response = requests.get(
        SEARCH_URL,
        params={"k": search_term},
        headers=HEADERS,
        timeout=20,
    )
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")
    offers: list[dict[str, str]] = []

    for result in soup.select('div[data-component-type="s-search-result"]'):
        title_tag = result.select_one("h2")
        link_tag = result.select_one('a[href*="/dp/"]')
        price_tag = result.select_one(".a-price")
        if not title_tag or not link_tag or not link_tag.get("href") or not price_tag:
            continue
        title = title_tag.get_text(" ", strip=True)
        whole = price_tag.select_one(".a-price-whole")
        fraction = price_tag.select_one(".a-price-fraction")
        symbol = price_tag.select_one(".a-price-symbol")
        if not whole:
            continue
        price = whole.get_text(strip=True).rstrip(".,")
        if fraction:
            price += f",{fraction.get_text(strip=True)}"
        if symbol:
            price += f" {symbol.get_text(strip=True)}"
        # Suchresultate können ähnlich benannte Staffeln enthalten. Alle
        # aussagekräftigen Wörter (einschließlich Zahlen) müssen daher vorkommen.
        term_words = {
            word.lower()
            for word in re.findall(r"\w+", search_term)
            if len(word) > 2 or word.isdigit()
        }
        title_words = set(re.findall(r"\w+", title.lower()))
        if not term_words.issubset(title_words):
            continue
        search_tokens = re.findall(r"\w+", search_term.lower())
        number_pairs = [
            (search_tokens[index - 1], token)
            for index, token in enumerate(search_tokens)
            if token.isdigit() and index > 0
        ]
        if any(
            not re.search(rf"\b{re.escape(word)}\s*[-:]?\s*{number}\b", title, re.IGNORECASE)
            for word, number in number_pairs
        ):
            continue
        link = urljoin("https://www.amazon.de", link_tag["href"])
        offers.append(
            {
                "Suchbegriff": search_term,
                "Produkt-Info": title,
                "Preis": price,
                "Verkäufer": seller_from_url(link),
                "Link": link,
                "Datum": date.today().isoformat(),
            }
        )
        if len(offers) >= limit:
            break
    return offers


def main() -> None:
    products = read_products(PRODUCTS_FILE)
    all_offers: list[dict[str, str]] = []

    for product in products:
        try:
            offers = find_offers(product)
            all_offers.extend(offers)
            print(f"{product}: {len(offers)} Angebot(e) gefunden")
        except requests.RequestException as error:
            print(f"{product}: Suche fehlgeschlagen ({error})")

    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    dataframe = pd.DataFrame(all_offers, columns=COLUMNS)
    dataframe.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
    print(f"Excel-Datei erstellt: {OUTPUT_FILE} ({len(dataframe)} Angebot(e))")


if __name__ == "__main__":
    main()
