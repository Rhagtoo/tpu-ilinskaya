# TPU Ильинская — scripts for final deck

## Установка

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Основные файлы

- `cadastre.py` — получить GeoJSON по кадастровому номеру через `rosreestr2coord`.
- `generate_maps_final.py` — пересобрать PNG-карты в `maps/`.
- `build_safe_final.py` — безопасно собрать базовую версию с картами из `TPU_Ilinskaya_reworked_final.pptx`.
- `fix_client_final_v2.py` — применить последние клиентские правки к `TPU_Ilinskaya_client_final.pptx` и получить `TPU_Ilinskaya_client_final_v2.pptx`.

## Воспроизведение финальной версии, которую сдавали последней

Положи в корень проекта файл:

```text
TPU_Ilinskaya_client_final.pptx
```

Затем:

```bash
python fix_client_final_v2.py \
  --input TPU_Ilinskaya_client_final.pptx \
  --output TPU_Ilinskaya_client_final_v2.pptx
```

Скрипт сам создаст временную папку `final_assets/`, извлечёт нужную картинку из PPTX, уберёт ошибочный маркер МКАД, пересоберёт разделители, титульный и контактный слайды.

## Пересборка карт и базовой версии

```bash
python generate_maps_final.py
python build_safe_final.py
```

Для этого рядом должны лежать:

```text
TPU_Ilinskaya_reworked_final.pptx
maps/
```

`generate_maps_final.py` создаёт `maps/` сам, но ему нужен доступ к OSM-тайлам. Если сети нет, карты не соберутся.

## Примечание по кадастру

Подтверждённая площадь участка: `206 700 м²` / `20,6 га`.
Файл `50_11_0050603_423.geojson` оставлен в репозитории как зафиксированный источник геометрии.
