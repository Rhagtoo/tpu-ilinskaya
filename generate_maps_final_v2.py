#!/usr/bin/env python3
"""
Генератор карт ТПУ Ильинская — v3.

Рендерит 7 статичных PNG-карт на OSM-тайлах + competitors_bg.png:
  maps/location_map.png — Об участке и местоположении
  maps/transport_map.png — Транспортная доступность
  maps/education_map.png — Образование и ДОУ
  maps/sport_map.png — Спорт и рекреация
  maps/trade_map.png — Торговля
  maps/medical_map.png — Медицина
  maps/competitors_bg.png — Конкурентное окружение (нумерованные маркеры)

Исправления v3 (относительно исходного v2):
 - Medical: zoom=11, центр (37.380, 55.752) — участок не уходит за край
 - Education: центр (37.400, 55.740) — лучший баланс участок/POI
 - Маркер «Участок»: r=11 + белое кольцо 4px для контраста
 - Шрифты: подписи 13px, заголовок 17px
 - Легенда: увеличенные вертикальные отступы
 - competitors_bg: нумерованные маркеры, цветовая легенда по классу

Зависимости:
    pip install staticmap pillow

Запуск:
    python generate_maps_final_v2.py
    # или только одну сцену:
    python generate_maps_final_v2.py --scene transport
    python generate_maps_final_v2.py --scene competitors
"""

from __future__ import annotations

import argparse
import math
import os
from dataclasses import dataclass
from typing import List, Tuple

from PIL import Image, ImageDraw, ImageFont
from staticmap import StaticMap

# =============================================================================
# CONFIG
# =============================================================================

OUTPUT_DIR = "maps"
WIDTH = 1200
MAP_H = 760
LEGEND_H = 110
HEIGHT = MAP_H + LEGEND_H
TILE_SIZE = 256

PLOT_AREA_HA = 20.6
PLOT_AREA_M2 = 206_700
CADASTRE_NUMBER = "50:11:0050603:423"

# Цвета
PLOT_COLOR = "#0D2137"
PLOT_OUTLINE = "#E74C3C"
TEXT_COLOR = "#1A2230"
MUTED_TEXT = "#6B7280"
WHITE = "#FFFFFF"
LABEL_BG = (255, 255, 255, 245)
LABEL_STROKE = "#C4C9D4"
ARROW_COLOR = "#0D2137"

POI_COLORS = [
    "#0D2137",
    "#2980B9",
    "#27AE60",
    "#E67E22",
    "#8E44AD",
    "#16A085",
    "#C0392B",
    "#2471A3",
    "#1E8449",
]

# =============================================================================
# GEO DATA
# =============================================================================

PLOT_POLYGON = [
    (37.2961114, 55.8029439), (37.2964606, 55.8020636), (37.2949235, 55.8018305),
    (37.2947308, 55.8022345), (37.2939885, 55.8020969), (37.2933638, 55.8021912),
    (37.2921096, 55.8016759), (37.2910849, 55.8010451), (37.2929445, 55.7983965),
    (37.2931575, 55.7980092), (37.2932150, 55.7971518), (37.2931318, 55.7963975),
    (37.2941164, 55.7966316), (37.2940644, 55.7966957), (37.2943066, 55.7967581),
    (37.2943618, 55.7966899), (37.2981301, 55.7975861), (37.2976478, 55.7995734),
    (37.2978250, 55.8011691), (37.2975658, 55.8025655), (37.2986215, 55.8033792),
    (37.2961114, 55.8029439),
]

PLOT_CENTER = (55.8000, 37.2949)  # lat, lon

COMPETITORS_POIS = [
    # (name, lat, lon)
    ("ТПУ Ильинская", 55.800, 37.295),       # 1 — целевой проект
    ("СберСити", 55.734, 37.478),             # 2 — бизнес
    ("Станиславский", 55.795, 37.330),        # 3 — бизнес
    ("Строгино 360", 55.800, 37.410),         # 4 — комфорт
    ("Квартал Строгино", 55.798, 37.403),     # 5 — комфорт
    ("Мыс", 55.720, 37.300),                  # 6 — бизнес/премиум
    ("Резиденции Сколково", 55.698, 37.359),  # 7 — бизнес-
    ("Сити Бэй", 55.750, 37.490),             # 8 — комфорт+
    ("Родина Переделкино", 55.645, 37.368),   # 9 — бизнес-
]

# Цвета маркеров конкурентов: 1=target(navy), 2,3,6,7,9=бизнес, 4,5,8=комфорт
COMP_COLORS = {
    1: ("#0D2137", True),    # Целевой проект
    2: ("#1E40AF", False),
    3: ("#1E40AF", False),
    4: ("#2563EB", False),
    5: ("#2563EB", False),
    6: ("#1E40AF", False),
    7: ("#1E40AF", False),
    8: ("#2563EB", False),
    9: ("#1E40AF", False),
}

# =============================================================================
# FONT HELPERS
# =============================================================================

def _load_font(size: int, bold: bool = False):
    candidates: List[str] = []
    if bold:
        candidates += [
            "/usr/share/fonts/truetype/noto/NotoSans-Bold.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/Library/Fonts/Arial Bold.ttf",
            "C:/Windows/Fonts/arialbd.ttf",
        ]
    candidates += [
        "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/Library/Fonts/Arial.ttf",
        "C:/Windows/Fonts/arial.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            return ImageFont.truetype(path, size=size)
    return ImageFont.load_default()


FONT_10 = _load_font(10)
FONT_11 = _load_font(11)
FONT_13 = _load_font(13)
FONT_13_B = _load_font(13, bold=True)
FONT_14_B = _load_font(14, bold=True)
FONT_17_B = _load_font(17, bold=True)

# =============================================================================
# SCENES
# =============================================================================

@dataclass
class Scene:
    name: str
    title: str
    pois: List[Tuple[str, float, float]]  # name, lat, lon
    center: Tuple[float, float]           # lon, lat
    zoom: int
    label_all: bool = True
    show_mkad_callout: bool = False


SCENES: List[Scene] = [
    Scene(
        name="location",
        title="Об участке и местоположении",
        pois=[("Участок", *PLOT_CENTER)],
        center=(37.2955, 55.8002),
        zoom=14,
    ),
    Scene(
        name="transport",
        title="Транспортная доступность",
        pois=[
            ("Участок", *PLOT_CENTER),
            ("Строгино", 55.804, 37.413),
            ("Москва-Сити", 55.749, 37.540),
        ],
        center=(37.405, 55.785),
        zoom=12,
        show_mkad_callout=True,
    ),
    Scene(
        name="education",
        title="Образование и ДОУ",
        pois=[
            ("Участок", *PLOT_CENTER),
            ("Сколково", 55.698, 37.359),
            ("МГИМО", 55.672, 37.487),
            ("Физтех-лицей", 55.753, 37.523),
            ("СберСити", 55.734, 37.478),
            ("Павловская гимн.", 55.755, 37.338),
        ],
        center=(37.400, 55.740),  # FIX v3: баланс участок/POI
        zoom=11,
    ),
    Scene(
        name="sport",
        title="Спорт и рекреация",
        pois=[
            ("Участок", *PLOT_CENTER),
            ("Серебряный Бор", 55.779, 37.418),
            ("Мещерский парк", 55.712, 37.413),
            ("Строгинская пойма", 55.804, 37.413),
            ("Крылатское", 55.757, 37.429),
        ],
        center=(37.370, 55.765),
        zoom=12,
    ),
    Scene(
        name="trade",
        title="Торговля",
        pois=[
            ("Участок", *PLOT_CENTER),
            ("Vegas Crocus", 55.821, 37.389),
            ("Твой Дом", 55.815, 37.407),
            ("METRO", 55.812, 37.382),
            ("Глобус", 55.792, 37.392),
        ],
        center=(37.355, 55.807),
        zoom=13,
    ),
    Scene(
        name="medical",
        title="Медицина",
        pois=[
            ("Участок", *PLOT_CENTER),
            ("Госпиталь Лапино", 55.671, 37.309),
            ("Кластер Сколково", 55.698, 37.359),
            ("Медицина СберСити", 55.734, 37.478),
        ],
        center=(37.380, 55.752),  # FIX v3: zoom=11, участок не уходит за край
        zoom=11,
    ),
]

# =============================================================================
# PROJECTION
# =============================================================================

def _mercator_world_px(lon: float, lat: float, zoom: int) -> Tuple[float, float]:
    lat = max(min(lat, 85.05112878), -85.05112878)
    sin_lat = math.sin(math.radians(lat))
    scale = TILE_SIZE * (2 ** zoom)
    x = (lon + 180.0) / 360.0 * scale
    y = (0.5 - math.log((1 + sin_lat) / (1 - sin_lat)) / (4 * math.pi)) * scale
    return x, y


def _project(lon: float, lat: float,
             center: Tuple[float, float], zoom: int) -> Tuple[int, int]:
    cx, cy = _mercator_world_px(center[0], center[1], zoom)
    px, py = _mercator_world_px(lon, lat, zoom)
    return int(round(WIDTH / 2 + (px - cx))), int(round(MAP_H / 2 + (py - cy)))


def _in_view(x: int, y: int, margin: int = 30) -> bool:
    return -margin <= x <= WIDTH + margin and -margin <= y <= MAP_H + margin

# =============================================================================
# DRAWING PRIMITIVES
# =============================================================================

def _text_bbox(draw: ImageDraw.ImageDraw,
               pos: Tuple[int, int], text: str, font) -> Tuple[int, int, int, int]:
    return draw.textbbox(pos, text, font=font)


def _draw_label(draw: ImageDraw.ImageDraw, x: int, y: int,
                text: str, preferred: str = "right") -> None:
    is_primary = text in {"Участок", "ТПУ Ильинская"}
    font = FONT_13_B if is_primary else FONT_13
    pad_x, pad_y = 7, 4

    if preferred == "left":
        tx, ty = x - 14, y - 12
        bbox0 = _text_bbox(draw, (0, 0), text, font)
        tx -= (bbox0[2] - bbox0[0])
    elif preferred == "above":
        bbox0 = _text_bbox(draw, (0, 0), text, font)
        tx = x - (bbox0[2] - bbox0[0]) // 2
        ty = y - 36
    else:
        tx, ty = x + 14, y - 12

    bbox = _text_bbox(draw, (tx, ty), text, font)
    rect = (bbox[0] - pad_x, bbox[1] - pad_y, bbox[2] + pad_x, bbox[3] + pad_y)

    # Clip to view
    if rect[2] > WIDTH - 8:
        dx = rect[2] - WIDTH + 8
        tx -= dx; rect = (rect[0]-dx, rect[1], rect[2]-dx, rect[3])
    if rect[0] < 8:
        dx = 8 - rect[0]
        tx += dx; rect = (rect[0]+dx, rect[1], rect[2]+dx, rect[3])
    if rect[1] < 70:
        dy = 70 - rect[1]
        ty += dy; rect = (rect[0], rect[1]+dy, rect[2], rect[3]+dy)
    if rect[3] > MAP_H - 8:
        dy = rect[3] - MAP_H + 8
        ty -= dy; rect = (rect[0], rect[1]-dy, rect[2], rect[3]-dy)

    draw.rounded_rectangle(rect, radius=6, fill=LABEL_BG, outline=LABEL_STROKE, width=1)
    draw.text((tx, ty), text, fill=TEXT_COLOR, font=font)


def _draw_marker(draw: ImageDraw.ImageDraw, x: int, y: int,
                 color: str, is_plot: bool = False) -> None:
    if is_plot:
        # White halo → color fill → dark border (r=11, ring=4px)
        draw.ellipse([x-15, y-15, x+15, y+15], fill=WHITE)
        draw.ellipse([x-11, y-11, x+11, y+11], fill=color,
                     outline="#0A1929", width=2)
    else:
        draw.ellipse([x-10, y-10, x+10, y+10], fill=WHITE)
        draw.ellipse([x-7, y-7, x+7, y+7], fill=color,
                     outline="#111827", width=1)


def _draw_mkad_callout(img: Image.Image) -> None:
    draw = ImageDraw.Draw(img, "RGBA")
    box = (WIDTH - 345, 88, WIDTH - 30, 152)
    draw.rounded_rectangle(
        box, radius=12,
        fill=(255, 255, 255, 245),
        outline=(13, 33, 55, 255), width=2,
    )
    draw.text((box[0] + 18, box[1] + 10), "Транспортный ориентир",
              fill=TEXT_COLOR, font=FONT_13)
    draw.text((box[0] + 18, box[1] + 30), "До МКАД ~5 км",
              fill=ARROW_COLOR, font=FONT_17_B)


def _draw_plot_polygon(img: Image.Image,
                       center: Tuple[float, float], zoom: int) -> Image.Image:
    rgba = img.convert("RGBA")
    overlay = Image.new("RGBA", (WIDTH, MAP_H), (255, 255, 255, 0))
    od = ImageDraw.Draw(overlay)
    poly_px = [_project(lon, lat, center, zoom) for lon, lat in PLOT_POLYGON]
    if len(poly_px) >= 3:
        od.polygon(poly_px, fill=(231, 76, 60, 35))
        od.line(poly_px, fill=(231, 76, 60, 255), width=3, joint="curve")
    return Image.alpha_composite(rgba, overlay).convert("RGB")


def _draw_pois(img: Image.Image, scene: Scene) -> Image.Image:
    draw = ImageDraw.Draw(img, "RGBA")
    for i, (name, lat, lon) in enumerate(scene.pois):
        x, y = _project(lon, lat, scene.center, scene.zoom)
        if not _in_view(x, y):
            continue
        is_plot = (i == 0) or name in {"Участок", "ТПУ Ильинская"}
        color = PLOT_COLOR if is_plot else POI_COLORS[i % len(POI_COLORS)]
        _draw_marker(draw, x, y, color, is_plot=is_plot)
        if scene.label_all or is_plot:
            preferred = (
                "above" if (is_plot and scene.name in {"education", "medical"})
                else "right"
            )
            _draw_label(draw, x, y, name, preferred=preferred)
    return img


def _draw_scene_header(img: Image.Image, title: str) -> None:
    overlay = Image.new("RGBA", (WIDTH, MAP_H), (255, 255, 255, 0))
    od = ImageDraw.Draw(overlay)
    bbox = od.textbbox((0, 0), title, font=FONT_17_B)
    text_w = bbox[2] - bbox[0]
    box_w = text_w + 36
    od.rounded_rectangle(
        (18, 16, 18 + box_w, 16 + 44),
        radius=10,
        fill=(255, 255, 255, 248),
        outline=(196, 201, 212, 255),
        width=1,
    )
    od.text((36, 27), title, fill=TEXT_COLOR, font=FONT_17_B)
    img_rgba = img.convert("RGBA")
    img_rgba.alpha_composite(overlay)
    img.paste(img_rgba.convert("RGB"))


def _draw_legend(base_map: Image.Image, scene: Scene) -> Image.Image:
    canvas = Image.new("RGB", (WIDTH, HEIGHT), WHITE)
    canvas.paste(base_map, (0, 0))
    draw = ImageDraw.Draw(canvas)

    draw.line([(0, MAP_H), (WIDTH, MAP_H)], fill="#D1D5DB", width=1)
    draw.text((20, MAP_H + 14), "Обозначения:", fill=TEXT_COLOR, font=FONT_13_B)

    x, y = 132, MAP_H + 13
    for i, (name, _, _) in enumerate(scene.pois):
        color = PLOT_COLOR if (i == 0 or name == "Участок") else POI_COLORS[i % len(POI_COLORS)]
        r = 7 if i == 0 else 5
        draw.ellipse([x, y + 4, x + r*2, y + 4 + r*2], fill=color, outline="#111827", width=1)
        draw.text((x + r*2 + 8, y + 1), name, fill=TEXT_COLOR, font=FONT_11)
        w = draw.textlength(name, font=FONT_11)
        x += int(w) + r*2 + 8 + 38
        if x > WIDTH - 220:
            x = 132
            y += 24

    y2 = MAP_H + 74
    x2 = 20
    draw.line([x2, y2 + 9, x2 + 28, y2 + 9], fill=PLOT_OUTLINE, width=3)
    draw.text(
        (x2 + 38, y2),
        f"Кадастровый участок {CADASTRE_NUMBER}: {PLOT_AREA_HA:.1f} га",
        fill="#993333", font=FONT_11,
    )
    draw.text((WIDTH - 250, y2), "© OpenStreetMap, Росреестр", fill=MUTED_TEXT, font=FONT_11)
    return canvas

# =============================================================================
# GENERATION
# =============================================================================

def render_osm_base(center: Tuple[float, float], zoom: int) -> Image.Image:
    sm = StaticMap(WIDTH, MAP_H)
    return sm.render(center=center, zoom=zoom).convert("RGB")


def generate_scene_map(scene: Scene, output_path: str) -> str:
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    img = render_osm_base(scene.center, scene.zoom)
    img = _draw_plot_polygon(img, scene.center, scene.zoom)
    img = _draw_pois(img, scene)
    if scene.show_mkad_callout:
        _draw_mkad_callout(img)
    _draw_scene_header(img, scene.title)
    canvas = _draw_legend(img, scene)
    canvas.save(output_path, quality=95)
    return output_path


def generate_competitors_bg(output_path: str) -> str:
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    center = (37.390, 55.745)
    zoom = 11
    img = render_osm_base(center, zoom)
    img = _draw_plot_polygon(img, center, zoom)

    draw = ImageDraw.Draw(img, "RGBA")
    legend_items = []

    for i, (name, lat, lon) in enumerate(COMPETITORS_POIS, start=1):
        x, y = _project(lon, lat, center, zoom)
        if not _in_view(x, y):
            continue
        color, is_target = COMP_COLORS.get(i, ("#2563EB", False))
        r = 14 if is_target else 11

        # Белое кольцо
        draw.ellipse([x-r-3, y-r-3, x+r+3, y+r+3], fill=WHITE)
        draw.ellipse([x-r, y-r, x+r, y+r], fill=color, outline="#111827", width=1)

        # Номер по центру
        num = str(i)
        bbox = _text_bbox(draw, (0, 0), num, FONT_13_B)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
        draw.text((x - tw / 2, y - th / 2 - 1), num, fill=WHITE, font=FONT_13_B)
        legend_items.append((i, name, color, is_target))

    _draw_scene_header(img, "Проекты конкурентного окружения")

    # ── Легенда ──────────────────────────────────────────────────────────────
    canvas = Image.new("RGB", (WIDTH, HEIGHT), (255, 255, 255))
    canvas.paste(img, (0, 0))
    d = ImageDraw.Draw(canvas)
    d.line([(0, MAP_H), (WIDTH, MAP_H)], fill="#D1D5DB", width=1)
    d.text((20, MAP_H + 12), "Обозначения", fill=TEXT_COLOR, font=FONT_13_B)

    # Нумерованные элементы легенды
    x, y = 140, MAP_H + 11
    for i, name, color, is_target in legend_items:
        r = 8 if is_target else 6
        d.ellipse([x, y + 4, x + r*2, y + 4 + r*2], fill=color,
                  outline="#111827", width=1)
        d.text((x + r - 3, y + 3), str(i), fill=WHITE, font=FONT_10)
        lbl = f"{i}. {name}"
        d.text((x + r*2 + 8, y + 1), lbl, fill=TEXT_COLOR, font=FONT_11)
        lw = d.textlength(lbl, font=FONT_11)
        x += int(lw) + r*2 + 8 + 22
        if x > WIDTH - 230:
            x = 140; y += 22

    # Цветовая легенда по классу
    cy2 = MAP_H + 76
    items_class = [
        ("#0D2137", "Целевой проект"),
        ("#1E40AF", "Бизнес-класс"),
        ("#2563EB", "Комфорт-класс"),
    ]
    cx2 = 20
    for color, label in items_class:
        d.ellipse([cx2, cy2 + 2, cx2 + 14, cy2 + 14], fill=color, outline="#111827", width=1)
        d.text((cx2 + 18, cy2), label, fill=TEXT_COLOR, font=FONT_11)
        cx2 += 18 + int(d.textlength(label, font=FONT_11)) + 24

    d.text((WIDTH - 250, cy2), "© OpenStreetMap, Росреестр", fill=MUTED_TEXT, font=FONT_11)

    canvas.save(output_path, quality=95)
    return output_path


# =============================================================================
# MAIN
# =============================================================================

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Генератор карт ТПУ Ильинская v3")
    parser.add_argument(
        "--scene", default=None,
        help="Имя сцены для генерации одной карты "
             "(location|transport|education|sport|trade|medical|competitors). "
             "Если не указано — генерируются все.")
    args = parser.parse_args()

    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    scene_map = {s.name: s for s in SCENES}

    if args.scene:
        if args.scene == "competitors":
            path = os.path.join(OUTPUT_DIR, "competitors_bg.png")
            generate_competitors_bg(path)
            print(f"  ✓ {path}")
        elif args.scene in scene_map:
            s = scene_map[args.scene]
            path = os.path.join(OUTPUT_DIR, f"{s.name}_map.png")
            generate_scene_map(s, path)
            print(f"  ✓ {path}  ({len(s.pois)} POI, zoom={s.zoom})")
        else:
            valid = list(scene_map.keys()) + ["competitors"]
            print(f"Неизвестная сцена {args.scene!r}. Доступные: {valid}")
            raise SystemExit(1)
    else:
        for scene in SCENES:
            path = os.path.join(OUTPUT_DIR, f"{scene.name}_map.png")
            generate_scene_map(scene, path)
            print(f"  ✓ {path}  ({len(scene.pois)} POI, zoom={scene.zoom})")

        comp_path = os.path.join(OUTPUT_DIR, "competitors_bg.png")
        generate_competitors_bg(comp_path)
        print(f"  ✓ {comp_path}  ({len(COMPETITORS_POIS)} POI)")

    print(f"\n✅ Готово — {len(SCENES)} карт + competitors_bg.png → '{OUTPUT_DIR}/'")
    print(f"   Участок {CADASTRE_NUMBER}, {PLOT_AREA_HA:.1f} га / {PLOT_AREA_M2:,} м²")


if __name__ == "__main__":
    main()
