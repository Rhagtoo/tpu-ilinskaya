# TPU Ильинская — Geo-Intelligence Report Generator

Автоматизированная сборка аналитических презентаций bnMAP.pro: OSM-карты + текстовая аналитика → PPTX.

**Участок:** 50:11:0050603:423 (центр: 55.8000, 37.2949)  
**Площадь:** ~20.6 га (206 700 м²)

---

## Файлы проекта

```
ilinskaya/
├── generate_maps_final.py       # Генератор 7 карт на OSM (PNG → maps/)
├── build_safe_final.py          # Сборка базовой PPTX (30 слайдов)
├── fix_client_final_v2.py       # Финальный патч (разделители, cover, позиционирование)
├── cadastre.py                  # Запрос геометрии из Росреестра
├── macbook_parser.py            # Парсер Avito (MacBook ремонт/запчасти)
├── requirements.txt             # Python-зависимости
├── README_BUILD.md              # Инструкция по сборке
├── 50_11_0050603_423.geojson    # Геометрия участка (22 точки)
├── maps/                        # Сгенерированные карты (7 PNG)
│   ├── location_map.png         #   Об участке и местоположении
│   ├── transport_map.png        #   Транспортная доступность
│   ├── education_map.png        #   Образование и ДОУ
│   ├── sport_map.png            #   Спорт и рекреация
│   ├── trade_map.png            #   Торговля
│   ├── medical_map.png          #   Медицина
│   └── competitors_bg.png       #   Конкурентное окружение
├── TPU_Ilinskaya_reworked_final.pptx  # Исходная презентация
├── TPU_Ilinskaya_final_safe.pptx      # Результат сборки
└── *.md (документация)
```

---

## Полный цикл

```bash
# Установка
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt

# Генерация карт
python generate_maps_final.py

# Сборка базовой презентации
python build_safe_final.py
# → TPU_Ilinskaya_final_safe.pptx (30 слайдов)

# Финальный патч (нужен TPU_Ilinskaya_client_final.pptx)
python fix_client_final_v2.py \
  --input TPU_Ilinskaya_client_final.pptx \
  --output TPU_Ilinskaya_client_final_v2.pptx
```

---

## Карты (6 сцен + competitors)

| # | Сцена | POI | Zoom |
|---|-------|-----|------|
| 1 | Об участке и местоположении | 1 | 14 |
| 2 | Транспортная доступность | 3 | 12 |
| 3 | Образование и ДОУ | 6 | 11 |
| 4 | Спорт и рекреация | 5 | 12 |
| 5 | Торговля | 5 | 13 |
| 6 | Медицина | 4 | 12 |
| 7 | Конкурентное окружение | 9 | 11 |

---

## Парсер Avito (macbook_parser.py)

Находит MacBook с признаками ремонта/неисправности на Avito. Использует `nodriver` (реальный Chrome, обходит анти-бот защиту).

```bash
pip install nodriver
python macbook_parser.py
# → avito_result_YYYYMMDD_HHMM.txt (название, цена, ссылка)
```
