"""Sammelt sichtbare Kaufangebote für die Begriffe aus ``produkte.txt``.

Unterstützt werden mehrere öffentlich sichtbare Shop-Suchen. Die Shops können
ihre Ergebnisse, Preise oder Schutzmechanismen jederzeit ändern; eine nicht
erreichbare Quelle wird übersprungen. Es werden höchstens 100 eindeutige
Angebote pro Lauf exportiert, damit der Abruf begrenzt bleibt.
"""

from __future__ import annotations

import re
import time
from datetime import date
from pathlib import Path
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


BASE_DIR = Path(__file__).resolve().parent
PRODUCTS_FILE = BASE_DIR / "produkte.txt"
OUTPUT_FILE = BASE_DIR / "ergebnisse" / "angebote.xlsx"
MAX_OFFERS = 100
MAX_PAGES_PER_AMAZON_STORE = 10
COLUMNS = ["Suchbegriff", "Produkt-Info", "Preis", "Verkäufer", "Link", "Datum"]
# Ein aktueller Browser-User-Agent; Captcha-/Bot-Schutzseiten werden nicht
# umgangen, sondern als nicht verfügbare Quelle behandelt.
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "de-DE,de;q=0.9,en-US;q=0.8,en;q=0.7",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
}


def read_products(path: Path) -> list[str]:
    """Liest nichtleere Suchbegriffe und ignoriert eine optionale Plaintext-Zeile."""
    if not path.exists():
        raise FileNotFoundError(f"Produktdatei nicht gefunden: {path}")
    return [
        line.strip()
        for line in path.read_text(encoding="utf-8").splitlines()
        if line.strip() and line.strip().lower() != "plaintext"
    ]


def create_session() -> requests.Session:
    """Erstellt eine Sitzung mit moderaten Wiederholungen für temporäre Fehler."""
    retry = Retry(
        total=2,
        backoff_factor=0.7,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=("GET",),
    )
    session = requests.Session()
    session.headers.update(HEADERS)
    session.mount("https://", HTTPAdapter(max_retries=retry))
    return session


def seller_from_url(url: str) -> str:
    host = urlparse(url).netloc.lower().removeprefix("www.")
    return host or "Unbekannter Verkäufer"


def search_variants(search_term: str) -> list[str]:
    """Erzeugt deutsche und englische Suchvarianten für internationale Shops."""
    translated = re.sub(r"\bstaffel\b", "Season", search_term, flags=re.IGNORECASE)
    return list(dict.fromkeys((search_term, translated)))


def title_matches_search(title: str, search_term: str) -> bool:
    """Prüft den Kern des Begriffs und Staffel-/Saisonnummern, nicht jedes Füllwort."""
    normalized_title = title.lower()
    title_words = set(re.findall(r"\w+", normalized_title))
    search_words = re.findall(r"\w+", search_term.lower())
    ignored = {"staffel", "season", "serie", "the", "complete"}
    core_words = {word for word in search_words if len(word) > 2 and word not in ignored}
    if not core_words.issubset(title_words):
        return False

    for index, token in enumerate(search_words):
        if not token.isdigit() or index == 0:
            continue
        label = search_words[index - 1]
        if label not in {"staffel", "season"}:
            continue
        # Akzeptiert "Season 2" sowie Boxen wie "Seasons 1-3", die Staffel 2 enthalten.
        match = re.search(
            rf"\b(?:staffel|season)s?\s*(\d+)\s*(?:[-–]\s*(\d+))?",
            normalized_title,
        )
        if not match:
            return False
        start, end = int(match.group(1)), int(match.group(2) or match.group(1))
        if not min(start, end) <= int(token) <= max(start, end):
            return False
    return True


def amazon_price(result: BeautifulSoup) -> str | None:
    price_tag = result.select_one(".a-price")
    if not price_tag:
        return None
    whole = price_tag.select_one(".a-price-whole")
    fraction = price_tag.select_one(".a-price-fraction")
    symbol = price_tag.select_one(".a-price-symbol")
    if not whole:
        return None
    price = whole.get_text(strip=True).rstrip(".,")
    if fraction:
        price += f",{fraction.get_text(strip=True)}"
    if symbol:
        price += f" {symbol.get_text(strip=True)}"
    return price


def canonical_amazon_link(store_url: str, href: str) -> str:
    """Entfernt Amazon-Empfehlungs- und Suchparameter aus derselben Produkt-URL."""
    link = urljoin(store_url, href)
    asin = re.search(r"/dp/([A-Z0-9]{10})", link, re.IGNORECASE)
    if asin:
        return f"{urlparse(store_url).scheme}://{urlparse(store_url).netloc}/dp/{asin.group(1)}"
    return link.split("?")[0]


def find_amazon_offers(
    session: requests.Session,
    store_url: str,
    search_term: str,
    remaining: int,
) -> list[dict[str, str]]:
    """Durchsucht mehrere Amazon-Ergebnisseiten bis zum gegebenen Restlimit."""
    offers: list[dict[str, str]] = []
    seen_links: set[str] = set()
    for query in search_variants(search_term):
        for page in range(1, MAX_PAGES_PER_AMAZON_STORE + 1):
            response = session.get(store_url, params={"k": query, "page": page}, timeout=25)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, "html.parser")
            results = soup.select('div[data-component-type="s-search-result"]')
            if not results:
                break
            page_count = 0
            for result in results:
                title_tag = result.select_one("h2")
                link_tag = result.select_one('a[href*="/dp/"]')
                if not title_tag or not link_tag or not link_tag.get("href"):
                    continue
                title = title_tag.get_text(" ", strip=True)
                if not title_matches_search(title, search_term):
                    continue
                price = amazon_price(result)
                if not price:
                    continue
                link = canonical_amazon_link(store_url, link_tag["href"])
                if link in seen_links:
                    continue
                seen_links.add(link)
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
                page_count += 1
                if len(offers) >= remaining:
                    return offers
            if page_count == 0:
                break
            time.sleep(0.5)
    return offers


def find_offers(search_term: str, remaining: int) -> list[dict[str, str]]:
    """Sammelt Angebote aus den konfigurierten, öffentlich erreichbaren Shops."""
    session = create_session()
    offers: list[dict[str, str]] = []
    sources = ("https://www.amazon.de/s", "https://www.amazon.com/s")
    for source in sources:
        if len(offers) >= remaining:
            break
        try:
            offers.extend(find_amazon_offers(session, source, search_term, remaining - len(offers)))
        except requests.RequestException as error:
            print(f"  {seller_from_url(source)} übersprungen: {error}")
    return offers


def main() -> None:
    all_offers: list[dict[str, str]] = []
    for product in read_products(PRODUCTS_FILE):
        remaining = MAX_OFFERS - len(all_offers)
        if remaining <= 0:
            break
        offers = find_offers(product, remaining)
        all_offers.extend(offers)
        print(f"{product}: {len(offers)} Angebot(e) gefunden")

    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    dataframe = pd.DataFrame(all_offers, columns=COLUMNS).drop_duplicates(subset=["Link"])
    dataframe.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
    print(f"Excel-Datei erstellt: {OUTPUT_FILE} ({len(dataframe)} Angebot(e), Limit {MAX_OFFERS})")


if __name__ == "__main__":
    main()
