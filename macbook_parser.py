#!/usr/bin/env python3
"""
Парсер MacBook (ремонт/запчасти) — Avito + Youla

Находит MacBook с признаками ремонта/неисправности:
  - Avito: [data-marker="item"] + stealth
  - Youla: a[href*="noutbuki"] с фильтром по ID

Фильтры: MacBook/M1-M5/год >= 2020, ремонтные ключевые слова,
исключает чехлы/зарядки/коробки.

Запуск:
    python macbook_parser.py

Результат: macbooks_result_YYYYMMDD_HHMM.xlsx
"""

from __future__ import annotations

import asyncio
import logging
import random
import re
from dataclasses import dataclass, asdict
from datetime import datetime
from urllib.parse import quote_plus, urljoin

import pandas as pd
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError
from playwright_stealth import Stealth

# =============================================================================
# НАСТРОЙКИ
# =============================================================================

SEARCH_QUERIES = [
    "MacBook M1 под ремонт",
    "MacBook M2 неисправен ремонт запчасти",
    "MacBook M3 под ремонт",
    "MacBook M4 нерабочий донор",
    "MacBook M1 ремонт",
    "MacBook M1 не включается",
    "MacBook M1 запчасти",
    "MacBook M2 ремонт",
    "MacBook M2 не включается",
    "MacBook M2 запчасти",
    "MacBook M3 ремонт",
    "MacBook M3 не включается",
    "MacBook M4 донор",
    "MacBook Air M1 разбит",
    "MacBook Pro M1 залит",
    "MacBook на запчасти",
]

HEADLESS = True               # False = видно окно браузера
MAX_AVITO_PER_QUERY = 20      # макс. карточек с Avito на запрос
MAX_YOULA_PER_QUERY = 15      # макс. карточек с Youla на запрос
ENABLE_DETAIL_PAGES = False   # True = открывать карточки (медленнее, риск капчи)

OUTPUT_FILE = f"macbooks_result_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

# =============================================================================
# КЛЮЧЕВЫЕ СЛОВА
# =============================================================================

KEYWORDS_REPAIR = [
    "ремонт", "неисправен", "неисправна", "не работает", "не включается",
    "донор", "на запчасти", "запчасти", "битый", "разбит", "разбитый",
    "экран", "матрица", "материнка", "плата", "залит", "залитый",
    "утопленник", "акб", "батарея", "шлейф", "клавиатура",
    "трещина", "полосы", "нет изображения",
]

NEGATIVE_KEYWORDS = [
    "чехол", "сумка", "зарядка", "адаптер", "блок питания",
    "наклейка", "подставка", "коробка", "мышь",
]

CHIP_RE = re.compile(r"\bM[1-5]\b", re.IGNORECASE)
YEAR_RE = re.compile(r"\b(202[0-9]|2030)\b")
PRICE_RE = re.compile(r"\d+")

MIN_YEAR = 2020

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
)


# =============================================================================
# МОДЕЛИ ДАННЫХ
# =============================================================================

@dataclass
class Listing:
    source: str
    title: str
    price_raw: str
    price_rub: int | None
    link: str
    location: str
    chip: str | None
    year: int | None
    query: str
    matched_keywords: str
    parsed_at: str


# =============================================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# =============================================================================

def clean_text(text: str | None) -> str:
    if not text:
        return ""
    return re.sub(r"\s+", " ", text).strip()


def parse_price(price_text: str | None) -> int | None:
    """'45 000 ₽' → 45000"""
    if not price_text:
        return None
    digits = PRICE_RE.findall(price_text.replace("\xa0", " "))
    if not digits:
        return None
    try:
        return int("".join(digits))
    except ValueError:
        return None


def extract_chip(text: str) -> str | None:
    m = CHIP_RE.search(text)
    return m.group(0).upper() if m else None


def extract_year(text: str) -> int | None:
    years = [int(y) for y in YEAR_RE.findall(text)]
    return max(years) if years else None


def get_matched_keywords(text: str) -> list[str]:
    lower = text.lower()
    return [kw for kw in KEYWORDS_REPAIR if kw.lower() in lower]


def is_relevant_macbook(text: str) -> tuple[bool, list[str]]:
    """
    Фильтр:
    - должен быть MacBook (или макбук)
    - должен быть чип M1-M5 ИЛИ год >= MIN_YEAR
    - должен быть признак ремонта/неисправности
    - отсекаем аксессуары
    """
    lower = text.lower()

    if "macbook" not in lower and "макбук" not in lower:
        return False, []

    if any(bad in lower for bad in NEGATIVE_KEYWORDS):
        return False, []

    chip = extract_chip(text)
    year = extract_year(text)

    if not chip and not (year and year >= MIN_YEAR):
        return False, []

    matched = get_matched_keywords(text)
    if not matched:
        return False, []

    return True, matched


async def safe_inner_text(element) -> str:
    if not element:
        return ""
    try:
        return clean_text(await element.inner_text())
    except Exception:
        return ""


async def safe_attr(element, attr: str) -> str:
    if not element:
        return ""
    try:
        return (await element.get_attribute(attr)) or ""
    except Exception:
        return ""


async def random_pause(min_sec: float = 1.5, max_sec: float = 3.5):
    await asyncio.sleep(random.uniform(min_sec, max_sec))


# =============================================================================
# ПАРСЕРЫ
# =============================================================================

async def scrape_avito(page, query: str) -> list[Listing]:
    """Парсер Avito с использованием Playwright Stealth."""
    results: list[Listing] = []
    url = f"https://www.avito.ru/all/noutbuki?q={quote_plus(query)}&s=104"

    logging.info("Avito: %s", query)
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=45000)
        await page.wait_for_timeout(4000)

        # Пробуем несколько селекторов (Avito меняет вёрстку)
        items = await page.query_selector_all('[data-marker="item"]')
        if not items:
            items = await page.query_selector_all(".iva-item-root")
        if not items:
            items = await page.query_selector_all('div[class*="iva-item"]')
        if not items:
            items = await page.query_selector_all('a[href*="/item"]')

        logging.info("Avito: найдено %d карточек", len(items))

        for item in items[:MAX_AVITO_PER_QUERY]:
            try:
                card_text = clean_text(await item.inner_text())
                if len(card_text) < 20:
                    continue

                # Заголовок
                title = ""
                for ts in ['[data-marker*="title"]', "h3", '[class*="title"]', '[itemprop="name"]']:
                    e = await item.query_selector(ts)
                    if e:
                        title = await safe_inner_text(e)
                        break
                if not title:
                    # fallback: текст до цены
                    parts = card_text.split("₽")
                    title = parts[0].strip()[-120:] if len(parts) > 1 else card_text[:120]

                # Цена
                price_raw = ""
                for ps in ['[data-marker*="price"]', '[itemprop="price"]', '[class*="price"]']:
                    e = await item.query_selector(ps)
                    if e:
                        price_raw = await safe_inner_text(e)
                        break
                if not price_raw:
                    pm = re.search(r"(\d[\d\s]*)\s*[₽]", card_text[:200])
                    price_raw = pm.group(0).strip() if pm else "Цена не указана"

                # Ссылка
                href = ""
                for a_el in await item.query_selector_all("a"):
                    h = await safe_attr(a_el, "href")
                    if h and "/item" in h:
                        href = h
                        break
                if not href:
                    href = await safe_attr(item, "href")
                link = urljoin("https://www.avito.ru", href) if href else ""

                full_text = f"{title} {card_text}"
                is_ok, matched = is_relevant_macbook(full_text)
                if not is_ok:
                    continue

                results.append(Listing(
                    source="Avito",
                    title=title[:150],
                    price_raw=price_raw,
                    price_rub=parse_price(price_raw),
                    link=link,
                    location="",
                    chip=extract_chip(full_text),
                    year=extract_year(full_text),
                    query=query,
                    matched_keywords=", ".join(matched),
                    parsed_at=datetime.now().isoformat(timespec="seconds"),
                ))

            except Exception as e:
                logging.debug("Avito: ошибка карточки: %s", e)

    except PlaywrightTimeoutError:
        logging.warning("Avito: таймаут загрузки страницы")
    except Exception as e:
        logging.error("Avito: ошибка запроса '%s': %s", query, e)

    return results


async def scrape_youla(page, query: str) -> list[Listing]:
    """Парсер Youla — ищет MacBook среди ноутбуков."""
    results: list[Listing] = []
    url = f"https://youla.ru/all/kompyutery/noutbuki?q={quote_plus(query)}"

    logging.info("Youla: %s", query)
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=30000)
        await page.wait_for_timeout(5000)

        # Youla: товары — это ссылки вида /moskva/.../noutbuki/XXXX-XXXX
        all_links = await page.query_selector_all('a[href*="noutbuki"]')
        items = [a for a in all_links
                 if re.search(r"[a-f0-9]{20,}", await safe_attr(a, "href") or "")]

        logging.info("Youla: найдено %d товаров (из %d ссылок)", len(items), len(all_links))

        for item in items[:MAX_YOULA_PER_QUERY]:
            try:
                card_text = clean_text(await item.inner_text())
                if len(card_text) < 15:
                    continue

                href = await safe_attr(item, "href")
                link = urljoin("https://youla.ru", href) if href else ""

                is_ok, matched = is_relevant_macbook(card_text)
                if not is_ok:
                    continue

                # Извлекаем цену и заголовок
                price_m = re.search(r"([\d\s]+)\s*[₽]", card_text)
                price_raw = price_m.group(0).strip() if price_m else "Цена не указана"
                title = card_text.split("₽")[-1].strip() if "₽" in card_text else card_text[:120]

                results.append(Listing(
                    source="Youla",
                    title=title[:150],
                    price_raw=price_raw,
                    price_rub=parse_price(price_raw),
                    link=link,
                    location="",
                    chip=extract_chip(card_text),
                    year=extract_year(card_text),
                    query=query,
                    matched_keywords=", ".join(matched),
                    parsed_at=datetime.now().isoformat(timespec="seconds"),
                ))

            except Exception as e:
                logging.debug("Youla: ошибка карточки: %s", e)

    except Exception as e:
        logging.error("Youla: ошибка запроса '%s': %s", query, e)

    return results


# =============================================================================
# ОБРАБОТКА РЕЗУЛЬТАТОВ
# =============================================================================

def deduplicate(listings: list[Listing]) -> list[Listing]:
    """Дедупликация по ссылке или source+title+price."""
    seen = set()
    unique = []
    for item in listings:
        key = item.link or f"{item.source}|{item.title}|{item.price_raw}"
        if key in seen:
            continue
        seen.add(key)
        unique.append(item)
    return unique


def save_results(listings: list[Listing]):
    """Сохранение в Excel."""
    if not listings:
        logging.info("Ничего не найдено")
        return

    rows = [asdict(x) for x in listings]
    df = pd.DataFrame(rows)

    columns = [
        "source", "title", "price_raw", "price_rub",
        "chip", "year", "location", "matched_keywords",
        "link", "query", "parsed_at",
    ]
    df = df[[col for col in columns if col in df.columns]]

    if "price_rub" in df.columns:
        df = df.sort_values(
            by=["price_rub", "source"],
            ascending=[True, True],
            na_position="last",
        )

    df.to_excel(OUTPUT_FILE, index=False)
    logging.info("✅ Сохранено %s объявлений → %s", len(df), OUTPUT_FILE)


# =============================================================================
# MAIN
# =============================================================================

async def main():
    stealth = Stealth()
    all_results: list[Listing] = []

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=HEADLESS,
            args=[
                "--no-sandbox",
                "--disable-setuid-sandbox",
                "--disable-blink-features=AutomationControlled",
            ],
        )
        context = await browser.new_context(
            viewport={"width": 1920, "height": 1080},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/134.0.0.0 Safari/537.36"
            ),
            locale="ru-RU",
            timezone_id="Europe/Moscow",
        )

        for i, query in enumerate(SEARCH_QUERIES, 1):
            logging.info("=== Запрос %d/%d ===", i, len(SEARCH_QUERIES))

            # Youla
            page_yl = await context.new_page()
            await stealth.apply_stealth_async(page_yl)
            try:
                r = await scrape_youla(page_yl, query)
                all_results.extend(r)
            except Exception as e:
                logging.error("Youla: %s", e)
            finally:
                await page_yl.close()

            await random_pause(3, 6)

            # Avito
            page_av = await context.new_page()
            await stealth.apply_stealth_async(page_av)
            try:
                r = await scrape_avito(page_av, query)
                all_results.extend(r)
            except Exception as e:
                logging.error("Avito: %s", e)
            finally:
                await page_av.close()

            await random_pause(4, 8)

        await context.close()
        await browser.close()

    unique = deduplicate(all_results)
    save_results(unique)


if __name__ == "__main__":
    asyncio.run(main())
