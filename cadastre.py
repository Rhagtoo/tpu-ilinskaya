#!/usr/bin/env python3
"""
Получение геометрии и данных участка по кадастровому номеру.
Использует официальное API Росреестра (через rosreestr2coord).

Установка: pip install rosreestr2coord
Запуск:    python cadastre.py 50:11:0050603:423
           python cadastre.py 50:11:0050603:423 --output result.geojson
"""

from __future__ import annotations
import argparse, json, os, sys
from rosreestr2coord.parser import Area


def get_cadastre(code: str, output_dir: str = "."):
    """
    Запрашивает геометрию и данные участка по кадастровому номеру.
    Возвращает dict с ключами: geojson, attrs, center, area_ha.
    """
    print(f"Запрос данных для {code}...")
    area = Area(code, use_cache=False, coord_out="EPSG:4326")

    # GeoJSON с геометрией
    geojson_str = area.to_geojson_poly()
    geojson = json.loads(geojson_str) if isinstance(geojson_str, str) else geojson_str

    # Атрибуты (площадь, стоимость, категория и т.д.)
    attrs = {}
    props = geojson.get("properties", {}).get("options", {})
    if props:
        attrs = {
            "cad_num": props.get("cad_num"),
            "area": props.get("declared_area"),
            "cost_value": props.get("cost_value"),
            "cost_index": props.get("cost_index"),
            "category": geojson["properties"].get("categoryName"),
        }

    # Координаты и центр
    coords = geojson["geometry"]["coordinates"][0]
    center_lon = sum(c[0] for c in coords) / len(coords)
    center_lat = sum(c[1] for c in coords) / len(coords)

    # Сохраняем
    os.makedirs(output_dir, exist_ok=True)
    name = code.replace(":", "_")
    geojson_path = os.path.join(output_dir, f"{name}.geojson")

    with open(geojson_path, "w") as f:
        json.dump(geojson, f, indent=2, ensure_ascii=False)
    print(f"  GeoJSON: {geojson_path}")

    return {
        "code": code,
        "center": (center_lat, center_lon),
        "points": len(coords),
        "attrs": attrs,
        "geojson_path": geojson_path,
        "coords": coords,
    }


def main():
    parser = argparse.ArgumentParser(description="Получение геометрии участка по кадастровому номеру")
    parser.add_argument("code", help="Кадастровый номер (например 50:11:0050603:423)")
    parser.add_argument("-o", "--output", default=".", help="Папка для сохранения GeoJSON")
    args = parser.parse_args()

    try:
        result = get_cadastre(args.code, args.output)
        print(f"\n{'='*50}")
        print(f"  Кадастровый номер: {result['code']}")
        print(f"  Центр: {result['center'][0]:.6f}, {result['center'][1]:.6f}")
        print(f"  Точек в полигоне: {result['points']}")
        for k, v in result["attrs"].items():
            if v is not None:
                if "cost" in k and isinstance(v, (int, float)):
                    print(f"  {k}: {v:,.0f}")
                else:
                    print(f"  {k}: {v}")
        print(f"  Файл: {result['geojson_path']}")
    except Exception as e:
        print(f"Ошибка: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
