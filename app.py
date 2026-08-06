from __future__ import annotations

import html
import math
import re
import unicodedata
from pathlib import Path

import folium
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from folium import FeatureGroup
from folium.plugins import Fullscreen, MeasureControl
from shapely.geometry import Point, Polygon, MultiPolygon
from shapely.ops import unary_union

# =========================================================
# AYÇA SİVAS ECZANE GRUP HARİTASI
# VERSION : V2.0
# DATE    : 06.08.2026
#
# CHANGELOG
# ---------------------------------------------------------
# V2.0
# - Grup Excel bağımlılığı tamamen kaldırıldı.
# - A1-D4 grupları doğrudan app.py içine gömüldü.
# - ZEREN yalnızca B4, İREM yalnızca C1 olarak tanımlandı.
# - Büyük pasta dilimleri ve büyük kesikli çemberler kaldırıldı.
# - Her grup sınırı, kendi eczanelerinin çevresinde küçük yerel
#   adalar halinde çizilir.
# - Sınır yarıçapı, en yakın farklı grup eczanesine göre otomatik
#   küçültülür; başka grubun eczanesine taşmaz.
# - Eczane adı ve grup yalnızca fareyle üzerine gelince görünür.
# =========================================================

st.set_page_config(
    page_title="Sivas Eczane Grup Haritası",
    page_icon="💊",
    layout="wide",
)

BASE_DIR = Path(__file__).resolve().parent
ECZANE_FILE_NAME = "nöbet-merkez tutan eczaneler(20260806-122152).xlsx"

MAIN_COLORS = {
    "A": "#E53935",
    "B": "#1E88E5",
    "C": "#43A047",
    "D": "#FB8C00",
}

SUBGROUP_DASH = {
    "1": None,
    "2": "7,5",
    "3": "3,5",
    "4": "10,5",
}

# =========================================================
# GRUP TANIMLARI
# Grup değişikliği için yalnızca bu bölümü düzenleyin.
# Bir eczane yalnızca tek bir grupta bulunmalıdır.
# =========================================================
GROUPS: dict[str, list[str]] = {
    "A1": [
        "BESTE", "HİKMET", "GÜNEŞ", "KARACA", "KOÇAK",
        "ALTINAY", "SELİN", "ÖZDEMİR",
    ],
    "A2": [
        "AYDOĞAN", "DERMAN", "DOLUNAY", "SEVGİ", "SERDAR",
        "SEVİNÇ", "ZEHRA", "ÇİĞDEM", "ÖRNEKOL",
    ],
    "A3": [
        "AKASYA", "ANIL", "BAĞDAT", "EĞRİKÖPRÜ", "GÖĞEBAKAN",
        "IŞIN", "NASUHOGLU", "ÜNİVERSİTE", "ESRAKAYA",
    ],
    "A4": [
        "DÖRTYOL", "GÜMÜŞ", "KURUGÖL", "NUR", "PAPATYA",
        "GÖKHAN", "SÜHA", "YENİEMEK",
    ],

    "B1": [
        "ADA", "AKER", "EBRU", "AYBÜKE", "KENT",
        "ÇAĞAN", "ÖZKAYNAK", "İSTASYON",
    ],
    "B2": [
        "ANADOLU", "BAHAR", "ELİF", "FATİH", "MAVİ",
        "SELÇUK", "TUĞBA", "YUNUSEMRE",
    ],
    "B3": [
        "ECE", "EKİCİ", "GÜLERSİN", "KALP", "İRFAN",
        "SELMA", "VATAN", "İLKER",
    ],
    "B4": [
        "AKSU", "ASLAN", "DUYGU", "ZEREN", "IŞIK",
        "KAĞAN", "MERAN", "YILDIRIM",
    ],

    "C1": [
        "BENGİSU", "AKIN", "BAŞAK", "AYDIN", "SİVAS",
        "İREM", "VERESELİÖMÜR", "AYKUT",
    ],
    "C2": [
        "ALPEREN", "BUKET", "DOĞA", "DOĞU", "DUMAN",
        "ERTUĞRUL", "GÖKÇE", "YENİŞEHİR", "DEMET",
    ],
    "C3": [
        "EKEN", "ERGÜN", "LALEZAR", "FURKAN", "IHLAMUR",
        "SAĞLIK", "TUĞRA", "ŞEYDA",
    ],
    "C4": [
        "DOĞANAY", "ALİBABA", "TÜLAY", "BERKAY", "YAPRAK",
        "İÇTEN", "RUMEYSA", "UĞUR", "MEYDAN",
    ],

    "D1": [
        "CEREN", "EREN", "ESRA", "TUĞUT", "YÖRÜKOĞLU",
        "ÇARŞI", "ÇOLAKOĞLU", "ŞENYURT", "KEPENEK",
    ],
    "D2": [
        "AKYOL", "FERHAT", "LOKMAN", "İSTANBUL",
        "MEVLANA", "MURAT", "ÜNAL", "GÜL",
    ],
    "D3": [
        "DOĞRUYOL", "ERSİN", "GÜLDEŞ", "HAKAN", "HALE",
        "KÜBRA", "SUBAŞI", "VEFA", "BEYZA",
    ],
    "D4": [
        "ARIKAN", "CAN", "KILIÇKAYA", "MEHMETAKİF",
        "SIHHAT", "TARIK", "VİTAMİN", "ÇETİNKAYA",
    ],
}


def normalize_name(value: object) -> str:
    """Türkçe karakter ve küçük yazım farklarını güvenli eşleştirir."""
    if pd.isna(value):
        return ""

    text = unicodedata.normalize("NFKC", str(value)).strip().upper()
    text = text.translate(
        str.maketrans(
            {
                "Ç": "C",
                "Ğ": "G",
                "İ": "I",
                "Ö": "O",
                "Ş": "S",
                "Ü": "U",
            }
        )
    )
    text = re.sub(r"\s+", "", text)
    text = re.sub(r"[^0-9A-Z]", "", text)

    # Koordinat dosyasında "ECZANESİ" son eki varsa eşleşmeyi bozmasın.
    if text.endswith("ECZANESI"):
        text = text[:-8]

    return text


def build_group_map() -> dict[str, str]:
    """GROUPS sözlüğünü tekil eczane -> grup haritasına çevirir."""
    group_map: dict[str, str] = {}

    for group_name, pharmacy_names in GROUPS.items():
        if not re.fullmatch(r"[ABCD][1-4]", group_name):
            raise ValueError(f"Geçersiz grup adı: {group_name}")

        for pharmacy_name in pharmacy_names:
            key = normalize_name(pharmacy_name)
            if not key:
                continue

            if key in group_map:
                raise ValueError(
                    f"'{pharmacy_name}' iki farklı grupta tanımlı: "
                    f"{group_map[key]} ve {group_name}"
                )

            group_map[key] = group_name

    # Kritik kayıtları açıkça doğrula.
    if group_map.get(normalize_name("ZEREN")) != "B4":
        raise ValueError("ZEREN mutlaka B4 grubunda olmalıdır.")

    if group_map.get(normalize_name("İREM")) != "C1":
        raise ValueError("İREM mutlaka C1 grubunda olmalıdır.")

    return group_map


def locate_pharmacy_file() -> Path | None:
    """Önce tam adı, ardından repo içindeki uygun koordinat Excel'ini bulur."""
    exact_path = BASE_DIR / ECZANE_FILE_NAME
    if exact_path.exists():
        return exact_path

    for candidate in sorted(BASE_DIR.glob("*.xlsx")):
        name_key = normalize_name(candidate.name)
        if "NOBETMERKEZTUTANECZANELER" in name_key:
            return candidate

    return None


def parse_coordinate(value: object) -> tuple[float, float] | None:
    if pd.isna(value):
        return None

    text = str(value).strip().replace(";", ",")
    match = re.search(
        r"(-?\d{1,3}(?:[\.,]\d+)?)\s*,\s*(-?\d{1,3}(?:[\.,]\d+)?)",
        text,
    )
    if not match:
        return None

    try:
        latitude = float(match.group(1).replace(",", "."))
        longitude = float(match.group(2).replace(",", "."))
    except ValueError:
        return None

    if not (-90 <= latitude <= 90 and -180 <= longitude <= 180):
        return None

    return latitude, longitude


@st.cache_data(show_spinner=False)
def read_pharmacies(path: str, file_version: int) -> pd.DataFrame:
    """
    Koordinat Excel'ini okur.
    file_version parametresi dosya değişince Streamlit önbelleğini yeniler.
    """
    del file_version

    raw = pd.read_excel(path, header=None, engine="openpyxl")
    records: list[dict[str, object]] = []

    for _, row in raw.iterrows():
        if len(row) < 2:
            continue

        name = "" if pd.isna(row.iloc[0]) else str(row.iloc[0]).strip()
        coordinate = parse_coordinate(row.iloc[1])

        if not name or coordinate is None:
            continue

        key = normalize_name(name)
        if key in {"ECZANE", "ECZANEADI", "AD"}:
            continue

        records.append(
            {
                "Eczane": name,
                "Anahtar": key,
                "Latitude": coordinate[0],
                "Longitude": coordinate[1],
            }
        )

    pharmacies = pd.DataFrame(records)
    if pharmacies.empty:
        raise ValueError(
            "Koordinat Excel'inde geçerli eczane ve koordinat bulunamadı."
        )

    duplicate_mask = pharmacies.duplicated(subset=["Anahtar"], keep=False)
    if duplicate_mask.any():
        duplicate_names = sorted(
            pharmacies.loc[duplicate_mask, "Eczane"].astype(str).unique()
        )
        raise ValueError(
            "Koordinat dosyasında yinelenen eczaneler var: "
            + ", ".join(duplicate_names)
        )

    return pharmacies.reset_index(drop=True)


def latlon_to_xy(
    latitude: float,
    longitude: float,
    reference_latitude: float,
    reference_longitude: float,
) -> tuple[float, float]:
    """Enlem-boylamı yerel metre koordinatlarına çevirir."""
    x = (
        (longitude - reference_longitude)
        * 111_320.0
        * math.cos(math.radians(reference_latitude))
    )
    y = (latitude - reference_latitude) * 111_320.0
    return x, y


def xy_to_latlon(
    x: float,
    y: float,
    reference_latitude: float,
    reference_longitude: float,
) -> tuple[float, float]:
    """Yerel metre koordinatlarını enlem-boylama geri çevirir."""
    latitude = reference_latitude + (y / 111_320.0)
    longitude = reference_longitude + (
        x
        / (
            111_320.0
            * max(math.cos(math.radians(reference_latitude)), 0.2)
        )
    )
    return latitude, longitude


def calculate_local_radius(
    row_index: int,
    points_xy: list[tuple[float, float]],
    groups: list[str],
) -> float:
    """
    Her eczane için sınır yarıçapı hesaplar.

    - En yakın farklı grup eczanesinin mesafesinin %42'sini geçmez.
    - Minimum 22 metre, maksimum 95 metre kullanır.
    - Böylece büyük çemberler oluşmaz ve başka grup noktasına ulaşılmaz.
    """
    x1, y1 = points_xy[row_index]
    current_group = groups[row_index]

    different_group_distances: list[float] = []

    for other_index, (x2, y2) in enumerate(points_xy):
        if other_index == row_index:
            continue
        if groups[other_index] == current_group:
            continue

        distance = math.hypot(x2 - x1, y2 - y1)
        different_group_distances.append(distance)

    if not different_group_distances:
        return 65.0

    nearest_other_group = min(different_group_distances)
    return max(22.0, min(95.0, nearest_other_group * 0.42))


def geometry_to_latlon(
    geometry: Polygon | MultiPolygon,
    reference_latitude: float,
    reference_longitude: float,
) -> list[list[tuple[float, float]]]:
    polygons: list[Polygon]

    if isinstance(geometry, Polygon):
        polygons = [geometry]
    elif isinstance(geometry, MultiPolygon):
        polygons = list(geometry.geoms)
    else:
        return []

    result: list[list[tuple[float, float]]] = []

    for polygon in polygons:
        if polygon.is_empty:
            continue

        coordinates: list[tuple[float, float]] = []
        for x, y in polygon.exterior.coords:
            coordinates.append(
                xy_to_latlon(
                    x,
                    y,
                    reference_latitude,
                    reference_longitude,
                )
            )

        if coordinates:
            result.append(coordinates)

    return result


def add_group_boundaries(map_obj: folium.Map, df: pd.DataFrame) -> None:
    """
    Her alt grup için küçük yerel adalar oluşturur.

    Aynı gruptaki yakın eczanelerin alanları birleşebilir.
    Uzak eczaneler dev bir sınırla birbirine bağlanmaz.
    """
    assigned_df = df.dropna(subset=["Grup"]).copy()
    if assigned_df.empty:
        return

    reference_latitude = float(assigned_df["Latitude"].median())
    reference_longitude = float(assigned_df["Longitude"].median())

    points_xy = [
        latlon_to_xy(
            float(row["Latitude"]),
            float(row["Longitude"]),
            reference_latitude,
            reference_longitude,
        )
        for _, row in assigned_df.iterrows()
    ]
    groups = assigned_df["Grup"].astype(str).tolist()

    radii = [
        calculate_local_radius(index, points_xy, groups)
        for index in range(len(points_xy))
    ]

    assigned_df = assigned_df.reset_index(drop=True)
    boundary_layer = FeatureGroup(name="Alt grup sınırları", show=True)

    for group_name, group_indices in assigned_df.groupby("Grup").groups.items():
        group_name = str(group_name)
        main_group = group_name[0]
        subgroup_number = group_name[1]
        color = MAIN_COLORS.get(main_group, "#616161")

        local_areas = [
            Point(points_xy[index]).buffer(radii[index], resolution=16)
            for index in group_indices
        ]

        geometry = unary_union(local_areas).buffer(0)
        geometry = geometry.simplify(2.0, preserve_topology=True)

        polygon_sets = geometry_to_latlon(
            geometry,
            reference_latitude,
            reference_longitude,
        )

        for polygon_coordinates in polygon_sets:
            folium.Polygon(
                locations=polygon_coordinates,
                color=color,
                weight=2.2,
                opacity=0.9,
                dash_array=SUBGROUP_DASH.get(subgroup_number),
                fill=True,
                fill_color=color,
                fill_opacity=0.055,
                tooltip=f"{group_name} sınırı",
            ).add_to(boundary_layer)

    boundary_layer.add_to(map_obj)


def add_pharmacy_markers(map_obj: folium.Map, df: pd.DataFrame) -> None:
    pharmacy_layer = FeatureGroup(name="Eczaneler", show=True)

    for _, row in df.iterrows():
        group_value = row.get("Grup")
        group_text = (
            str(group_value)
            if pd.notna(group_value)
            else "Grupsuz"
        )

        main_group = (
            group_text[0]
            if group_text and group_text[0] in MAIN_COLORS
            else ""
        )
        color = MAIN_COLORS.get(main_group, "#757575")

        tooltip_html = (
            '<div style="font-size:14px; line-height:1.35;">'
            f'<b>{html.escape(str(row["Eczane"]))}</b><br>'
            f'Grup: <b>{html.escape(group_text)}</b>'
            "</div>"
        )

        folium.CircleMarker(
            location=[
                float(row["Latitude"]),
                float(row["Longitude"]),
            ],
            radius=5.8,
            color="#FFFFFF",
            weight=1.5,
            fill=True,
            fill_color=color,
            fill_opacity=0.96,
            tooltip=folium.Tooltip(
                tooltip_html,
                sticky=True,
                direction="top",
            ),
        ).add_to(pharmacy_layer)

    pharmacy_layer.add_to(map_obj)


def build_map(df: pd.DataFrame) -> folium.Map:
    center = [
        float(df["Latitude"].median()),
        float(df["Longitude"].median()),
    ]

    map_obj = folium.Map(
        location=center,
        zoom_start=13,
        tiles=None,
        control_scale=True,
        prefer_canvas=True,
    )

    folium.TileLayer(
        tiles="CartoDB positron",
        name="Sade harita",
        control=True,
        show=True,
    ).add_to(map_obj)

    folium.TileLayer(
        tiles="OpenStreetMap",
        name="Detaylı harita",
        control=True,
        show=False,
    ).add_to(map_obj)

    add_group_boundaries(map_obj, df)
    add_pharmacy_markers(map_obj, df)

    Fullscreen(
        position="topright",
        title="Tam ekran",
        title_cancel="Tam ekrandan çık",
    ).add_to(map_obj)

    MeasureControl(
        position="topright",
        primary_length_unit="meters",
    ).add_to(map_obj)

    folium.LayerControl(
        collapsed=False,
        position="topright",
    ).add_to(map_obj)

    map_obj.fit_bounds(
        [
            [
                float(df["Latitude"].min()),
                float(df["Longitude"].min()),
            ],
            [
                float(df["Latitude"].max()),
                float(df["Longitude"].max()),
            ],
        ],
        padding=(25, 25),
    )

    return map_obj


# =========================================================
# UYGULAMA
# =========================================================
st.title("Sivas Eczane Grup Haritası")
st.caption(
    "V2.0 — Gruplar doğrudan app.py içindedir. "
    "Eczane adı ve grubu yalnızca fareyle üzerine gelince görünür."
)

pharmacy_path = locate_pharmacy_file()

if pharmacy_path is None:
    st.error(
        "Koordinat Excel dosyası GitHub reposunda bulunamadı. "
        "app.py ile aynı klasöre aşağıdaki dosyayı yükleyin:"
    )
    st.code(ECZANE_FILE_NAME)
    st.stop()

try:
    group_map = build_group_map()

    pharmacies = read_pharmacies(
        str(pharmacy_path),
        pharmacy_path.stat().st_mtime_ns,
    )
    pharmacies["Grup"] = pharmacies["Anahtar"].map(group_map)

    assigned_count = int(pharmacies["Grup"].notna().sum())
    missing_count = int(pharmacies["Grup"].isna().sum())

    col1, col2, col3 = st.columns(3)
    col1.metric("Toplam eczane", len(pharmacies))
    col2.metric("Grubu eşleşen", assigned_count)
    col3.metric("Grupsuz", missing_count)

    zeren_rows = pharmacies[
        pharmacies["Anahtar"] == normalize_name("ZEREN")
    ]
    if zeren_rows.empty:
        st.warning("Koordinat dosyasında ZEREN bulunamadı.")
    elif zeren_rows.iloc[0]["Grup"] != "B4":
        st.error("Kritik hata: ZEREN B4 olarak eşleşmedi.")
        st.stop()

    if missing_count:
        missing_names = ", ".join(
            pharmacies.loc[
                pharmacies["Grup"].isna(),
                "Eczane",
            ].astype(str).tolist()
        )
        with st.expander(
            f"Grubu eşleşmeyen {missing_count} eczaneyi göster"
        ):
            st.write(missing_names)

    pharmacy_map = build_map(pharmacies)
    components.html(
        pharmacy_map.get_root().render(),
        height=900,
        scrolling=False,
    )

except Exception as exc:
    st.exception(exc)
