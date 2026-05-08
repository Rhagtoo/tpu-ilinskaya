#!/usr/bin/env python3
"""
Парсер Avito — MacBook ремонт/запчасти
Выход: avito_result_YYYYMMDD_HHMM.txt (название, цена, ссылка)
"""
import asyncio, re, random, logging, os
from datetime import datetime
from urllib.parse import quote_plus
import nodriver as uc

QUERIES = [
    "MacBook M1 под ремонт", "MacBook M2 неисправен запчасти",
    "MacBook M3 под ремонт", "MacBook M4 донор",
    "MacBook M1 ремонт", "MacBook M1 не включается",
    "MacBook M1 запчасти", "MacBook M2 ремонт",
    "MacBook M2 не включается", "MacBook M3 ремонт",
    "MacBook на запчасти", "MacBook разбит донор",
]
MAX_ITEMS = 15
PROFILE_DIR = os.path.join(os.path.dirname(__file__) or ".", "chrome_profile")
OUTPUT_FILE = os.path.join(os.path.dirname(__file__) or ".",
                           f"avito_result_{datetime.now().strftime('%Y%m%d_%H%M')}.txt")

KW = ["ремонт","неисправен","не работает","не включается","донор","на запчасти",
      "запчасти","битый","разбит","экран","матрица","плата","залит","утопленник",
      "акб","батарея","клавиатура","трещина","полосы","нет изображения"]
NEG = ["чехол","сумка","зарядка","адаптер","блок питания","наклейка","подставка","коробка","мышь"]

logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(message)s")

async def main():
    os.makedirs(PROFILE_DIR, exist_ok=True)
    browser = await uc.start(headless=False, user_data_dir=PROFILE_DIR)
    logging.info("Chrome запущен\n")

    out = open(OUTPUT_FILE, "w", encoding="utf-8-sig")
    out.write("Цена\tНазвание\tСсылка\n")

    total = 0; seen = set()
    for i, q in enumerate(QUERIES, 1):
        logging.info("=== %d/%d: %s ===", i, len(QUERIES), q)
        try:
            tab = await browser.get(
                f"https://www.avito.ru/all/noutbuki?q={quote_plus(q)}&s=104")
            await asyncio.sleep(random.uniform(5, 8))

            items = await tab.select_all('[data-marker="item"]')
            if not items:
                items = await tab.select_all('.iva-item-root')

            for item in items[:MAX_ITEMS]:
                try:
                    text = item.text_all
                    if not text or len(text) < 20: continue
                    lower = text.lower()
                    if "macbook" not in lower and "макбук" not in lower: continue
                    if any(n in lower for n in NEG): continue
                    if not [k for k in KW if k.lower() in lower]: continue

                    # Ссылка: ищем ЛЮБОЙ <a> внутри карточки
                    link = ""
                    a = await item.query_selector('a')
                    if a:
                        href = (a.attrs or {}).get("href", "")
                        if href and href.startswith("/"):
                            link = "https://www.avito.ru" + href
                        elif href:
                            link = href

                    # Цена
                    price = "?"
                    price_el = await item.query_selector('[data-marker*="price"],[class*="price"]')
                    price_text = (price_el.text_all or "").strip() if price_el else ""
                    if not price_text:
                        for line in text.split("\n")[:10]:
                            pm = re.search(r"(\d[\d\s]+)\s*[₽]", line)
                            if pm: price = re.sub(r"\s","",pm.group(1)); break
                    else:
                        pm = re.search(r"([\d\s]+)\s*[₽]", price_text)
                        if pm: price = re.sub(r"\s","",pm.group(1))

                    # Заголовок
                    title = ""
                    title_el = await item.query_selector('[data-marker*="title"],h3,[class*="title"]')
                    if title_el: title = (title_el.text_all or "").strip()
                    if not title:
                        title = text.split("₽")[-1].strip()[:120] if "₽" in text else text[:120]
                    title = title.replace("\t"," ").replace("\n"," ")

                    key = link or title
                    if key in seen: continue
                    seen.add(key); total += 1

                    out.write(f"{price}\t{title}\t{link}\n")
                    out.flush()
                    print(f"  {price}₽ | {title[:80]}")
                except:
                    pass
        except Exception as e:
            logging.error("Ошибка: %s", e)
        await asyncio.sleep(random.uniform(4, 8))

    out.close()
    try: browser.stop()
    except: pass

    if total:
        print(f"\n✅ {total} объявлений → {OUTPUT_FILE}")
    else:
        print("\n❌ Ничего не найдено")

if __name__ == "__main__":
    asyncio.run(main())
