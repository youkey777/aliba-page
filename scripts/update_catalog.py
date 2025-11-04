# -*- coding: utf-8 -*-
"""Regenerate catalog data and update HTML sections based on the latest Excel spec."""
from __future__ import annotations

import html
import json
import re
import time
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from datetime import datetime, timedelta
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd
import requests

BASE_DIR = Path(__file__).resolve().parents[1]
EXCEL_PATH = Path("C:/Users/hfree/.claude/projects/20251027_aliba_mens_page/画像指定/カテゴリ別画像.xlsx")

PRODUCT_CACHE_PATH = BASE_DIR / "assets" / "product_cache.json"
PRODUCT_RECORDS_PATH = BASE_DIR / "assets" / "product_records.json"
COMPILED_PRODUCTS_PATH = BASE_DIR / "assets" / "compiled_products.json"
GROUPED_PRODUCTS_PATH = BASE_DIR / "assets" / "grouped_products.json"
PRICE_LIST_PATH = BASE_DIR / "assets" / "price_list.json"
PRICE_STATUS_PATH = BASE_DIR / "assets" / "price_status.json"
SPECIFIED_PRODUCTS_PATH = BASE_DIR / "assets" / "specified_products.json"
INDEX_HTML_PATH = BASE_DIR / "index.html"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36",
    "Accept-Language": "ja-JP,ja;q=0.9,en-US;q=0.8,en;q=0.7",
}

COLUMN_LOOKUP = {
    "\\u51fa\\u54c1\\u8005SKU": "sku",
    "ASIN 1": "asin",
    "\\u5546\\u54c1\\u540d": "name",
    "\\u4fa1\\u683c": "price_excel",
    "\\u30ab\\u30c6\\u30b4\\u30ea": "category",
    "\\u5728\\u5eab\\u6570": "stock",
}

CATEGORY_LOOKUP = {
    "\\u30ec\\u30c7\\u30a3\\u30fc\\u30b9\\u30d4\\u30a2\\u30b9": {"key": "ladies_earrings", "page": "ladies-earrings-page", "card": "item"},
    "\\u30cd\\u30c3\\u30af\\u30ec\\u30b9": {"key": "ladies_necklaces", "page": "ladies-necklaces-page", "card": "item"},
    "\\u30bb\\u30c3\\u30c8\\u30a2\\u30a4\\u30c6\\u30e0": {"key": "set_items", "page": "set-items-page", "card": "item"},
    "\\u30da\\u30a2\\u30cd\\u30c3\\u30af\\u30ec\\u30b9": {"key": "pair_necklaces", "page": "pair-necklaces-page", "card": "item"},
    "\\u30e1\\u30f3\\u30ba\\u30d4\\u30a2\\u30b9": {"key": "mens_earrings", "page": "mens-earrings-page", "card": "item"},
    "\\u30e1\\u30f3\\u30ba\\u30cd\\u30c3\\u30af\\u30ec\\u30b9": {"key": "mens_necklaces", "page": "mens-necklaces-page", "card": "item"},
    "\\u8ca1\\u5e03": {"key": "mens_wallets", "page": "mens-wallets-page", "card": "item"},
    "\\u30cd\\u30af\\u30bf\\u30a4\\u30d4\\u30f3": {"key": "mens_tiepins", "page": "mens-tiepins-page", "card": "item"},
    "\\u907a\\u9aa8\\u30cd\\u30c3\\u30af\\u30ec\\u30b9": {"key": "memorial_items", "page": "memorial-page", "card": "memorial"},
    "\\u907a\\u9aa8\\u6839\\u30af\\u30ec\\u30b9": {"key": "memorial_items", "page": "memorial-page", "card": "memorial"},
    "alivaluxe": {"key": "aliva_luxe", "page": "aliva-luxe-page", "card": "product"},
}

CATEGORY_META: Dict[str, Dict[str, str]] = {}
for _meta in CATEGORY_LOOKUP.values():
    CATEGORY_META.setdefault(_meta["key"], _meta)

REQUEST_TIMEOUT = 15
FETCH_DELAY = 1.2

YEN = "\u00a5"
CACHE_MAX_AGE_HOURS = 72


@dataclass
class Product:
    asin: str
    name: str
    price: str
    price_value: int
    image: str
    url: str

    @property
    def image_key(self) -> str:
        if "._" in self.image:
            return self.image.split("._")[0]
        return self.image


def rename_columns(df: pd.DataFrame) -> pd.DataFrame:
    renamed = {}
    for column in df.columns:
        encoded = column.encode("unicode_escape").decode("ascii")
        renamed[column] = COLUMN_LOOKUP.get(encoded, column)
    return df.rename(columns=renamed)


def load_excel() -> pd.DataFrame:
    df = pd.read_excel(EXCEL_PATH)
    return rename_columns(df)


def normalise_category(value) -> Optional[Dict[str, str]]:
    if pd.isna(value):
        return None
    encoded = str(value).encode("unicode_escape").decode("ascii")
    return CATEGORY_LOOKUP.get(encoded)


def parse_stock(value) -> float:
    if pd.isna(value):
        return 1.0
    if isinstance(value, str):
        s = value.strip()
        if not s:
            return 0.0
        lowered = s.lower()
        if lowered in {"na", "nan", "not answer", "notanswered", "not-answered"}:
            return 0.0
        if s in {"ノットアンサー", "なし"}:
            return 0.0
        s = s.replace(",", "")
        try:
            return float(s)
        except ValueError:
            return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def sanitize_price(price_str: Optional[str]) -> Optional[str]:
    if not price_str:
        return None
    value = price_str.replace("￥", YEN).replace("\uffe5", YEN)
    match = re.search(r"(\d[\d,]*)", value)
    if not match:
        return None
    digits = match.group(1).replace(",", "")
    try:
        number = int(digits)
    except ValueError:
        return None
    return f"{YEN}{number:,}"


def price_to_int(price_str: str) -> int:
    digits = re.sub(r"[^0-9]", "", price_str)
    return int(digits) if digits else 0


def fetch_amazon_product(asin: str) -> Optional[Dict[str, str]]:
    url = f"https://www.amazon.co.jp/dp/{asin}"
    try:
        response = requests.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT)
        if response.status_code != 200:
            return None
    except requests.RequestException:
        return None

    try:
        from bs4 import BeautifulSoup

        soup = BeautifulSoup(response.text, "html.parser")
    except Exception:
        return None

    price = None
    for span in soup.select("span.a-offscreen"):
        cleaned = sanitize_price(span.get_text(strip=True))
        if cleaned:
            price = cleaned
            break

    image = None
    img = soup.select_one("#landingImage")
    if img:
        image = img.get("data-old-hires") or img.get("src")
    if not image:
        og = soup.select_one('meta[property="og:image"]')
        if og:
            image = og.get("content")
    if not image:
        return None

    title_el = soup.select_one("#productTitle")
    title = title_el.get_text(strip=True) if title_el else None

    return {
        "asin": asin,
        "url": url,
        "title": title,
        "price": price,
        "image": image,
        "fetched_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
        "status_code": response.status_code,
    }


def is_entry_stale(entry: Dict[str, str]) -> bool:
    fetched_at = entry.get("fetched_at")
    if not fetched_at:
        return True
    try:
        fetched_time = datetime.strptime(fetched_at, "%Y-%m-%dT%H:%M:%S")
    except ValueError:
        return True
    return datetime.utcnow() - fetched_time >= timedelta(hours=CACHE_MAX_AGE_HOURS)


def update_product_cache(
    asins: Iterable[str],
    force_refresh: Optional[Iterable[str]] = None,
) -> Dict[str, Dict[str, str]]:
    if PRODUCT_CACHE_PATH.exists():
        cache = json.loads(PRODUCT_CACHE_PATH.read_text(encoding="utf-8"))
    else:
        cache = {}

    force_set = {asin.strip() for asin in force_refresh or [] if asin}

    updated = False
    for asin in asins:
        asin = asin.strip()
        if not asin:
            continue
        entry = cache.get(asin)
        if (
            asin not in force_set
            and entry
            and entry.get("price")
            and entry.get("image")
            and not is_entry_stale(entry)
        ):
            continue
        fetched = fetch_amazon_product(asin)
        if fetched:
            cache[asin] = fetched
            updated = True
            time.sleep(FETCH_DELAY)

    if updated:
        ordered = dict(sorted(cache.items()))
        PRODUCT_CACHE_PATH.write_text(json.dumps(ordered, ensure_ascii=False, indent=2), encoding="utf-8")
        return ordered
    return cache


def build_product(asin: str, name: str, cache: Dict[str, Dict[str, str]], fallback_price: Optional[float]) -> Optional[Product]:
    return build_product_with_fallbacks(
        asin=asin,
        name=name,
        cache=cache,
        fallback_price=fallback_price,
        fallback_image=None,
    )


def build_product_with_fallbacks(
    asin: str,
    name: Optional[str],
    cache: Dict[str, Dict[str, str]],
    fallback_price: Optional[float],
    fallback_image: Optional[str],
) -> Optional[Product]:
    entry = cache.get(asin, {})
    price = sanitize_price(entry.get("price"))
    if not price and fallback_price:
        price = f"{YEN}{int(fallback_price):,}"
    if not price:
        return None
    price_value = price_to_int(price)
    image = entry.get("image") or fallback_image
    if not image:
        return None
    url = entry.get("url", f"https://www.amazon.co.jp/dp/{asin}")
    product_name = name or entry.get("title") or asin
    return Product(
        asin=asin,
        name=product_name,
        price=price,
        price_value=price_value,
        image=image,
        url=url,
    )


def load_specified_products() -> Dict[str, List[Dict[str, object]]]:
    if not SPECIFIED_PRODUCTS_PATH.exists():
        return {}
    try:
        return json.loads(SPECIFIED_PRODUCTS_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}


def prepare_specified_products(
    specified: Dict[str, List[Dict[str, object]]],
    cache: Dict[str, Dict[str, str]],
) -> Dict[str, List[Dict[str, object]]]:
    prepared: Dict[str, List[Dict[str, object]]] = {}
    for key, entries in specified.items():
        section: List[Dict[str, object]] = []
        for entry in sorted(entries, key=lambda item: item.get("order", 0)):
            asin_raw = entry.get("asin")
            asin = str(asin_raw).strip() if asin_raw is not None else ""
            if not asin:
                continue
            try:
                order = int(entry.get("order", len(section) + 1))
            except (TypeError, ValueError):
                order = len(section) + 1
            product = build_product_with_fallbacks(
                asin=asin,
                name=entry.get("name"),
                cache=cache,
                fallback_price=None,
                fallback_image=entry.get("image_fallback"),
            )
            if not product:
                continue
            section.append(
                {
                    "order": order,
                    "display_name": entry.get("name") or product.name,
                    "product": product,
                }
            )
        if section:
            prepared[key] = section
    return prepared


def find_div_bounds(html_text: str, marker: str, occurrence: int = 1) -> Tuple[int, int, int]:
    position = -1
    search_start = 0
    for _ in range(occurrence):
        position = html_text.find(marker, search_start)
        if position == -1:
            raise ValueError(f"Marker {marker} not found")
        search_start = position + len(marker)
    content_start = position + len(marker)
    depth = 1
    idx = content_start
    while depth > 0:
        next_open = html_text.find("<div", idx)
        next_close = html_text.find("</div>", idx)
        if next_close == -1:
            raise ValueError("Unbalanced div structure")
        if next_open != -1 and next_open < next_close:
            depth += 1
            idx = next_open + 4
        else:
            depth -= 1
            idx = next_close + len("</div>")
    return position, content_start, idx


def detect_indent(html_text: str, position: int) -> str:
    newline_index = html_text.rfind("\n", 0, position)
    if newline_index == -1:
        return ""
    idx = newline_index + 1
    indent_chars: List[str] = []
    while idx < position and html_text[idx] in {" ", "\t"}:
        indent_chars.append(html_text[idx])
        idx += 1
    return "".join(indent_chars)


def replace_div_inner(html_text: str, class_name: str, blocks: List[str], occurrence: int = 1) -> str:
    marker = f'<div class="{class_name}">'
    start, content_start, end = find_div_bounds(html_text, marker, occurrence)
    indent = detect_indent(html_text, start)
    inner_indent = indent + "  "
    closing_start = end - len("</div>")

    if blocks:
        inner_lines: List[str] = []
        for block in blocks:
            for line in block.split("\n"):
                inner_lines.append(inner_indent + line)
        new_inner = "\n" + "\n".join(inner_lines) + "\n" + indent
    else:
        new_inner = "\n" + indent

    return html_text[:content_start] + new_inner + html_text[closing_start:]


def render_ranking_items(items: List[Dict[str, object]]) -> List[str]:
    cards: List[str] = []
    for item in items:
        product: Product = item["product"]  # type: ignore[assignment]
        display_name = str(item.get("display_name", product.name))
        order = item.get("order", "")
        lines = [
            '<div class="ranking-item">',
            f'  <div class="ranking-number">{html.escape(str(order))}</div>',
            f'  <a class="ranking-link" href="{html.escape(product.url)}" rel="noopener noreferrer" target="_blank">',
            '    <div class="ranking-image">',
            f'    <img alt="{html.escape(display_name)}" src="{html.escape(product.image)}"/>',
            "    </div>",
            '    <div class="ranking-info">',
            f'    <div class="ranking-name">{html.escape(display_name)}</div>',
            f'    <div class="ranking-price">{product.price}</div>',
            "    </div>",
            "  </a>",
            "</div>",
        ]
        cards.append("\n".join(lines))
    return cards


def render_all_items(items: List[Dict[str, object]]) -> List[str]:
    cards: List[str] = []
    for item in items:
        product: Product = item["product"]  # type: ignore[assignment]
        display_name = str(item.get("display_name", product.name))
        lines = [
            f'<a class="item-card" href="{html.escape(product.url)}" rel="noopener noreferrer" target="_blank">',
            '  <div class="item-image">',
            f'  <img alt="{html.escape(display_name)}" src="{html.escape(product.image)}"/>',
            "  </div>",
            '  <div class="item-info">',
            f'  <div class="item-name">{html.escape(display_name)}</div>',
            f'  <div class="item-price">{product.price}</div>',
            "  </div>",
            "</a>",
        ]
        cards.append("\n".join(lines))
    return cards


def group_and_sort(products: List[Product]) -> Tuple[List[Product], List[List[Product]]]:
    groups: Dict[str, List[Product]] = defaultdict(list)
    for product in products:
        groups[product.image_key].append(product)

    ordered: List[Tuple[int, List[Product]]] = []
    for items in groups.values():
        items.sort(key=lambda p: (-p.price_value, p.asin))
        ordered.append((max(p.price_value for p in items), items))

    ordered.sort(key=lambda entry: (-entry[0], entry[1][0].asin))

    flattened: List[Product] = []
    grouped: List[List[Product]] = []
    for _, items in ordered:
        grouped.append(items)
        flattened.extend(items)
    return flattened, grouped


def render_item_cards(products: List[Product], class_name: str = "item-card") -> List[str]:
    cards = []
    for product in products:
        href = html.escape(product.url)
        alt = html.escape(product.name)
        src = html.escape(product.image)
        lines = [
            f'<a class="{class_name}" href="{href}" rel="noopener noreferrer" target="_blank">',
            '  <div class="item-image">',
            f'  <img alt="{alt}" src="{src}"/>',
            '  </div>',
            '  <div class="item-info">',
            f'  <div class="item-price">{product.price}</div>',
            '  </div>',
            '</a>',
        ]
        cards.append("\n".join(lines))
    return cards


def render_product_cards(products: List[Product]) -> List[str]:
    cards = []
    for product in products:
        href = html.escape(product.url)
        alt = html.escape(product.name)
        src = html.escape(product.image)
        lines = [
            f'<a class="product-card" href="{href}" target="_blank">',
            '  <div class="product-image">',
            f'  <img alt="{alt}" src="{src}"/>',
            '  </div>',
            '  <div class="product-info">',
            f'  <div class="product-price">{product.price}</div>',
            '  </div>',
            '</a>',
        ]
        cards.append("\n".join(lines))
    return cards


def render_memorial_placeholders(count: int = 6) -> List[str]:
    cards = []
    for _ in range(count):
        lines = [
            '<div class="product-card coming-soon-card">',
            '  <div class="product-image">',
            '  <div class="coming-soon-icon" aria-hidden="true">&#8987;</div>',
            '  </div>',
            '  <div class="product-info">',
            '  <div class="coming-soon-label">Coming Soon</div>',
            '  </div>',
            '</div>',
        ]
        cards.append("\n".join(lines))
    return cards


def find_grid_bounds(html_text: str, page_id: str) -> Tuple[int, int, int]:
    id_marker = f'id="{page_id}"'
    id_pos = html_text.find(id_marker)
    if id_pos == -1:
        raise ValueError(f"Page {page_id} not found")
    div_start = html_text.rfind('<div', 0, id_pos)
    if div_start == -1:
        raise ValueError(f"Div start not found for {page_id}")
    grid_marker = '<div class="products-grid">'
    grid_start = html_text.find(grid_marker, div_start)
    if grid_start == -1:
        raise ValueError(f"products-grid not found for {page_id}")
    content_start = grid_start + len(grid_marker)

    depth = 1
    idx = content_start
    while depth > 0:
        next_open = html_text.find('<div', idx)
        next_close = html_text.find('</div>', idx)
        if next_close == -1:
            raise ValueError("Unbalanced div structure")
        if next_open != -1 and next_open < next_close:
            depth += 1
            idx = next_open + 4
        else:
            depth -= 1
            idx = next_close + len('</div>')
    grid_end = idx
    return grid_start, content_start, grid_end


def replace_products_grid(html_text: str, page_id: str, cards: List[str]) -> str:
    grid_start, content_start, grid_end = find_grid_bounds(html_text, page_id)
    inner = "\n" + "\n".join("  " + line for card in cards for line in card.split("\n")) + "\n"
    new_section = html_text[grid_start:content_start] + inner + '</div>'
    return html_text[:grid_start] + new_section + html_text[grid_end:]


def ensure_css_snippet(html_text: str) -> str:
    snippet = (
        "        .coming-soon-card {\n"
        "            background: rgba(255, 255, 255, 0.08);\n"
        "            border: 1px dashed rgba(255, 255, 255, 0.35);\n"
        "            border-radius: 16px;\n"
        "            padding: 40px 20px;\n"
        "            display: flex;\n"
        "            flex-direction: column;\n"
        "            align-items: center;\n"
        "            justify-content: center;\n"
        "            gap: 16px;\n"
        "            text-align: center;\n"
        "        }\n\n"
        "        .coming-soon-card .coming-soon-label {\n"
        "            font-size: 18px;\n"
        "            letter-spacing: 0.12em;\n"
        "            text-transform: uppercase;\n"
        "            color: var(--fg-primary);\n"
        "        }\n\n"
        "        .coming-soon-card .product-price {\n"
        "            font-size: 16px;\n"
        "            letter-spacing: 0.08em;\n"
        "            color: var(--fg-secondary);\n"
        "        }\n\n"
        "        .memorial-theme .coming-soon-card {\n"
        "            background: rgba(240, 244, 249, 0.35);\n"
        "            border-color: rgba(184, 197, 214, 0.5);\n"
        "        }\n\n"
        "        .memorial-theme .coming-soon-card .coming-soon-label {\n"
        "            color: #4B5563;\n"
        "        }\n\n"
        "        .memorial-theme .coming-soon-card .product-price {\n"
        "            color: #6B7280;\n"
        "        }\n"
    )
    if snippet in html_text:
        return html_text
    marker = "/* Memorial Theme - Elegant & Delicate Style (Light Silver) */"
    if marker not in html_text:
        return html_text
    return html_text.replace(marker, snippet + "\n        " + marker)


def main() -> None:
    df = load_excel()

    records = []
    for _, row in df.iterrows():
        records.append(
            {
                "asin": row.get("asin"),
                "name": row.get("name"),
                "price_excel": float(row.get("price_excel")) if not pd.isna(row.get("price_excel")) else None,
                "category_raw": row.get("category"),
                "stock": row.get("stock"),
            }
        )
    PRODUCT_RECORDS_PATH.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")

    specified_raw = load_specified_products()

    valid_entries: List[Tuple[Dict[str, str], str, pd.Series]] = []
    for _, row in df.iterrows():
        meta = normalise_category(row.get("category"))
        if not meta:
            continue
        if parse_stock(row.get("stock")) <= 0:
            continue
        asin = str(row.get("asin")).strip()
        if not asin:
            continue
        valid_entries.append((meta, asin, row))

    specified_asins = {
        str(entry.get("asin")).strip()
        for entries in specified_raw.values()
        for entry in entries
        if entry.get("asin")
    }
    fetch_targets = {asin for _, asin, _ in valid_entries if asin}
    fetch_targets.update({asin for asin in specified_asins if asin})

    aliva_luxe_asins = {asin for meta, asin, _ in valid_entries if meta["key"] == "aliva_luxe"}

    cache = update_product_cache(fetch_targets)

    compiled: Dict[str, List[Product]] = defaultdict(list)
    for meta, asin, row in valid_entries:
        product = build_product(asin, row.get("name"), cache, row.get("price_excel"))
        if product:
            compiled[meta["key"]].append(product)

    specified_prepared = prepare_specified_products(specified_raw, cache)

    compiled_serialisable: Dict[str, List[Dict[str, object]]] = {}
    grouped_serialisable: Dict[str, List[List[Dict[str, object]]]] = {}
    price_lists: Dict[str, List[str]] = {}
    price_status: Dict[str, List[Dict[str, object]]] = {}
    rendered_cards: Dict[str, List[str]] = {}

    for key, meta in CATEGORY_META.items():
        products = compiled.get(key, [])
        if meta["card"] == "memorial":
            compiled_serialisable[key] = []
            grouped_serialisable[key] = []
            price_lists[key] = []
            price_status[key] = []
            rendered_cards[key] = render_memorial_placeholders()
            continue
        if not products:
            continue
        flattened, grouped = group_and_sort(products)
        compiled_serialisable[key] = [p.__dict__ for p in flattened]
        grouped_serialisable[key] = [[p.__dict__ for p in group] for group in grouped]
        price_lists[key] = [p.price for p in flattened]
        price_status[key] = [
            {"asin": p.asin, "price": p.price, "has_price": True, "has_image": True}
            for p in flattened
        ]
        if meta["card"] == "product":
            rendered_cards[key] = render_product_cards(flattened)
        else:
            rendered_cards[key] = render_item_cards(flattened)

    for key, items in specified_prepared.items():
        price_status[f"specified_{key}"] = [
            {
                "asin": item["product"].asin,  # type: ignore[index]
                "price": item["product"].price,  # type: ignore[index]
                "has_price": True,
                "has_image": True,
            }
            for item in items
        ]

    COMPILED_PRODUCTS_PATH.write_text(json.dumps(compiled_serialisable, ensure_ascii=False, indent=2), encoding="utf-8")
    GROUPED_PRODUCTS_PATH.write_text(json.dumps(grouped_serialisable, ensure_ascii=False, indent=2), encoding="utf-8")
    PRICE_LIST_PATH.write_text(json.dumps(price_lists, ensure_ascii=False, indent=2), encoding="utf-8")
    PRICE_STATUS_PATH.write_text(json.dumps(price_status, ensure_ascii=False, indent=2), encoding="utf-8")

    html_text = INDEX_HTML_PATH.read_text(encoding="utf-8")
    for key, meta in CATEGORY_META.items():
        cards = rendered_cards.get(key)
        if not cards:
            continue
        try:
            html_text = replace_products_grid(html_text, meta["page"], cards)
        except ValueError:
            continue

    ladies_ranking = specified_prepared.get("ladies_ranking")
    if ladies_ranking:
        try:
            html_text = replace_div_inner(html_text, "ranking-scroll", render_ranking_items(ladies_ranking), occurrence=1)
        except ValueError:
            pass

    mens_ranking = specified_prepared.get("mens_ranking")
    if mens_ranking:
        try:
            html_text = replace_div_inner(html_text, "ranking-scroll", render_ranking_items(mens_ranking), occurrence=2)
        except ValueError:
            pass

    ladies_all = specified_prepared.get("ladies_all")
    if ladies_all:
        try:
            html_text = replace_div_inner(html_text, "items-grid", render_all_items(ladies_all), occurrence=1)
        except ValueError:
            pass

    mens_all = specified_prepared.get("mens_all")
    if mens_all:
        try:
            html_text = replace_div_inner(html_text, "items-grid", render_all_items(mens_all), occurrence=2)
        except ValueError:
            pass

    html_text = ensure_css_snippet(html_text)
    INDEX_HTML_PATH.write_text(html_text, encoding="utf-8")


if __name__ == "__main__":
    main()
