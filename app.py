afrom __future__ import annotations

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
from branca.element import MacroElement, Template
from shapely.geometry import LineString, Point, Polygon, MultiPolygon
from shapely.ops import unary_union

# =========================================================
# AYÇA SİVAS ECZANE GRUP HARİTASI
# VERSION : V3.0
# DATE    : 13.08.2026
#
# CHANGELOG
# ---------------------------------------------------------
# V2.2
# - A, B, C ve D ana grupları kendi renk ailesine ayrıldı.
# - Her alt grup 1'den 4'e koyudan açığa farklı bir ton kullanır.
# - Marker, sınır çizgisi ve dolgu aynı alt grup rengini kullanır.
#
# V2.1
# - Aynı alt gruptaki yakın eczaneler küme halinde sınır içine alınır.
# - Bir küme en az 3 eczaneden oluşur.
# - Yakında 5, 6, 7 veya daha fazla aynı alt grup eczanesi varsa
#   tamamı tek sınır içinde gösterilir.
# - Uzak kümeler dev bir alanla birbirine bağlanmaz.
# - Diğer grup eczanelerinin çevresi sınır geometrisinden çıkarılır.
#
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
ECZANE_FILE_NAME = "nöbet-merkez tutan eczaneler(20260813-054852).xlsx"

GROUP_COLORS = {
    # A grubu: mavi tonları
    "A1": "#0D47A1",
    "A2": "#1976D2",
    "A3": "#42A5F5",
    "A4": "#90CAF9",

    # B grubu: yeşil tonları
    "B1": "#1B5E20",
    "B2": "#388E3C",
    "B3": "#66BB6A",
    "B4": "#A5D6A7",

    # C grubu: turuncu tonları
    "C1": "#E65100",
    "C2": "#F57C00",
    "C3": "#FFB74D",
    "C4": "#FFE0B2",

    # D grubu: mor tonları
    "D1": "#4A148C",
    "D2": "#7B1FA2",
    "D3": "#BA68C8",
    "D4": "#E1BEE7",
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
        "BESTE", "HİKMET", "ÖRNEKOL", "KARACA", "KOÇAK",
        "ALTINAY", "SELİN", "ÖZDEMİR",
    ],
    "A2": [
        "AYDOĞAN", "DERMAN", "DOLUNAY", "SEVGİ", "SERDAR",
        "SEVİNÇ", "ZEHRA", "ÇİĞDEM", "GÜNEŞ",
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
        "ÇARŞI", "ÇOLAKOĞLU", "LOKMAN", "KEPENEK",
    ],
    "D2": [
        "AKYOL", "FERHAT", "ŞENYURT", "İSTANBUL",
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


def point_distance(
    first: tuple[float, float],
    second: tuple[float, float],
) -> float:
    return math.hypot(second[0] - first[0], second[1] - first[1])


def create_initial_clusters(
    indices: list[int],
    points_xy: list[tuple[float, float]],
    connection_distance_m: float = 350.0,
) -> list[list[int]]:
    """
    Aynı alt gruptaki eczaneleri yakınlık ağına göre kümeler.

    Zincirleme yakınlık geçerlidir: A, B'ye; B de C'ye yakınsa
    üçü aynı kümede değerlendirilir.
    """
    adjacency: dict[int, list[int]] = {index: [] for index in indices}

    for position, first_index in enumerate(indices):
        for second_index in indices[position + 1:]:
            distance = point_distance(
                points_xy[first_index],
                points_xy[second_index],
            )
            if distance <= connection_distance_m:
                adjacency[first_index].append(second_index)
                adjacency[second_index].append(first_index)

    clusters: list[list[int]] = []
    visited: set[int] = set()

    for start_index in indices:
        if start_index in visited:
            continue

        stack = [start_index]
        visited.add(start_index)
        cluster: list[int] = []

        while stack:
            current = stack.pop()
            cluster.append(current)

            for neighbour in adjacency[current]:
                if neighbour not in visited:
                    visited.add(neighbour)
                    stack.append(neighbour)

        clusters.append(sorted(cluster))

    return clusters


def cluster_distance(
    first_cluster: list[int],
    second_cluster: list[int],
    points_xy: list[tuple[float, float]],
) -> float:
    """İki küme arasındaki en yakın eczane mesafesini verir."""
    return min(
        point_distance(points_xy[first], points_xy[second])
        for first in first_cluster
        for second in second_cluster
    )


def enforce_minimum_cluster_size(
    clusters: list[list[int]],
    points_xy: list[tuple[float, float]],
    minimum_size: int = 3,
) -> list[list[int]]:
    """
    Üçten küçük kümeleri en yakın aynı alt grup kümesine birleştirir.

    Böylece her sınır mümkün olduğunda en az üç eczaneyi kapsar.
    Alt grubun toplam eczane sayısı üçten azsa mevcutların tamamı kullanılır.
    """
    clusters = [list(cluster) for cluster in clusters]

    while len(clusters) > 1:
        small_positions = [
            position
            for position, cluster in enumerate(clusters)
            if len(cluster) < minimum_size
        ]
        if not small_positions:
            break

        source_position = min(
            small_positions,
            key=lambda position: len(clusters[position]),
        )
        source_cluster = clusters[source_position]

        target_positions = [
            position
            for position in range(len(clusters))
            if position != source_position
        ]
        target_position = min(
            target_positions,
            key=lambda position: cluster_distance(
                source_cluster,
                clusters[position],
                points_xy,
            ),
        )

        merged = sorted(source_cluster + clusters[target_position])

        for position in sorted(
            [source_position, target_position],
            reverse=True,
        ):
            clusters.pop(position)
        clusters.append(merged)

    return sorted(clusters, key=lambda cluster: min(cluster))


def nearest_other_group_distance(
    point_index: int,
    points_xy: list[tuple[float, float]],
    groups: list[str],
) -> float:
    current_group = groups[point_index]
    distances = [
        point_distance(points_xy[point_index], points_xy[other_index])
        for other_index in range(len(points_xy))
        if other_index != point_index
        and groups[other_index] != current_group
    ]
    return min(distances) if distances else 200.0


def build_cluster_geometry(
    cluster: list[int],
    points_xy: list[tuple[float, float]],
    groups: list[str],
) -> Polygon | MultiPolygon:
    """
    Küme noktalarını ince koridorlarla bağlayarak sıkı bir sınır üretir.

    Dışbükey büyük alan kullanılmaz. Bu nedenle sınır boş bölgeleri
    gereksiz yere kaplamaz.
    """
    point_radii: dict[int, float] = {}
    for index in cluster:
        nearest_other = nearest_other_group_distance(
            index,
            points_xy,
            groups,
        )
        point_radii[index] = max(18.0, min(48.0, nearest_other * 0.32))

    geometry_parts = [
        Point(points_xy[index]).buffer(point_radii[index], resolution=18)
        for index in cluster
    ]

    # Minimum spanning tree: noktaları en kısa toplam bağlantıyla birleştirir.
    if len(cluster) >= 2:
        connected = {cluster[0]}
        remaining = set(cluster[1:])

        while remaining:
            first, second, distance = min(
                (
                    connected_index,
                    remaining_index,
                    point_distance(
                        points_xy[connected_index],
                        points_xy[remaining_index],
                    ),
                )
                for connected_index in connected
                for remaining_index in remaining
            )

            corridor_width = max(
                12.0,
                min(
                    30.0,
                    point_radii[first] * 0.65,
                    point_radii[second] * 0.65,
                ),
            )
            geometry_parts.append(
                LineString(
                    [points_xy[first], points_xy[second]]
                ).buffer(corridor_width, cap_style=1, join_style=1)
            )
            connected.add(second)
            remaining.remove(second)

    geometry = unary_union(geometry_parts).buffer(0)

    # Başka grupların eczane noktalarını sınırın dışında bırak.
    cluster_group = groups[cluster[0]]
    exclusion_areas = [
        Point(points_xy[index]).buffer(16.0, resolution=14)
        for index in range(len(points_xy))
        if groups[index] != cluster_group
    ]
    if exclusion_areas:
        geometry = geometry.difference(unary_union(exclusion_areas)).buffer(0)

    return geometry.simplify(1.5, preserve_topology=True)


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
        if polygon.is_empty or polygon.area < 80.0:
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



def add_group_boundaries(
    map_obj: folium.Map,
    full_df: pd.DataFrame,
    selected_groups: set[str],
) -> None:
    """
    Seçili alt grupların sınırlarını çizer.

    Önemli: Geometri hesabı tüm eczaneler üzerinden yapılır. Böylece bir alt grup
    filtrelendiğinde, görünmeyen diğer gruplar yokmuş gibi sınırlar genişlemez.
    """
    assigned_df = full_df.dropna(subset=["Grup"]).copy().reset_index(drop=True)
    if assigned_df.empty or not selected_groups:
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

    for group_name, group_indices_value in assigned_df.groupby("Grup").groups.items():
        group_name = str(group_name)
        if group_name not in selected_groups:
            continue

        group_indices = sorted(int(index) for index in group_indices_value)
        subgroup_number = group_name[1]
        color = GROUP_COLORS.get(group_name, "#616161")

        # Her alt grup ayrı layer: sağ üst Leaflet menüsünden de tek tek kapatılabilir.
        boundary_layer = FeatureGroup(
            name=f"{group_name} sınırı",
            show=True,
        )

        clusters = create_initial_clusters(
            group_indices,
            points_xy,
            connection_distance_m=350.0,
        )
        clusters = enforce_minimum_cluster_size(
            clusters,
            points_xy,
            minimum_size=3,
        )

        for cluster in clusters:
            geometry = build_cluster_geometry(
                cluster,
                points_xy,
                groups,
            )

            polygon_sets = geometry_to_latlon(
                geometry,
                reference_latitude,
                reference_longitude,
            )

            pharmacy_names = assigned_df.loc[cluster, "Eczane"].astype(str).tolist()
            tooltip_text = (
                f"{group_name} — {len(cluster)} eczane: "
                + ", ".join(pharmacy_names)
            )

            for polygon_coordinates in polygon_sets:
                folium.Polygon(
                    locations=polygon_coordinates,
                    color=color,
                    weight=2.3,
                    opacity=0.92,
                    dash_array=SUBGROUP_DASH.get(subgroup_number),
                    fill=True,
                    fill_color=color,
                    fill_opacity=0.06,
                    tooltip=tooltip_text,
                ).add_to(boundary_layer)

        boundary_layer.add_to(map_obj)


def add_pharmacy_markers(
    map_obj: folium.Map,
    df: pd.DataFrame,
    selected_groups: set[str],
) -> None:
    """Seçili alt grupların eczanelerini, her alt grup ayrı Leaflet katmanı olacak şekilde ekler."""
    if df.empty or not selected_groups:
        return

    for group_name in sorted(selected_groups):
        subset = df[df["Grup"] == group_name]
        if subset.empty:
            continue

        pharmacy_layer = FeatureGroup(
            name=f"{group_name} eczaneleri",
            show=True,
        )

        color = GROUP_COLORS.get(group_name, "#757575")

        for _, row in subset.iterrows():
            tooltip_html = (
                '<div style="font-size:14px; line-height:1.35;">'
                f'<b>{html.escape(str(row["Eczane"]))}</b><br>'
                f'Grup: <b>{html.escape(group_name)}</b>'
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


class DensityCircleControl(MacroElement):
    """Leaflet üzerinde yoğunluk çemberi ve canlı eczane sayacı."""

    def __init__(
        self,
        pharmacy_points: list[dict[str, object]],
        center_lat: float,
        center_lon: float,
        count_label: str,
    ):
        super().__init__()
        self._name = "DensityCircleControl"

        import json

        pharmacies_json = json.dumps(pharmacy_points, ensure_ascii=False)
        count_label_json = json.dumps(count_label, ensure_ascii=False)

        self._template = Template(
            r"""
{% macro script(this, kwargs) %}
(function() {
    const map = {{ this._parent.get_name() }};
    const pharmacies = {{ this.pharmacies_json | safe }};
    const total = pharmacies.length;
    const countLabel = {{ this.count_label_json | safe }};

    const startLatLng = L.latLng({{ this.center_lat }}, {{ this.center_lon }});

    const densityCircle = L.circle(startLatLng, {
        radius: 1000,
        color: '#C62828',
        weight: 3,
        opacity: 0.95,
        fillColor: '#EF5350',
        fillOpacity: 0.12,
        interactive: false
    }).addTo(map);

    const centerIcon = L.divIcon({
        className: '',
        html: '<div style="width:24px;height:24px;border-radius:50%;background:#C62828;border:4px solid white;box-shadow:0 1px 6px rgba(0,0,0,.45);cursor:move;"></div>',
        iconSize: [24, 24],
        iconAnchor: [12, 12]
    });

    const centerHandle = L.marker(startLatLng, {
        draggable: true,
        icon: centerIcon,
        zIndexOffset: 2000,
        title: 'Çember merkezini sürükle'
    }).addTo(map);

    const DensityControl = L.Control.extend({
        options: { position: 'topleft' },
        onAdd: function() {
            const div = L.DomUtil.create('div', 'ayca-density-panel');
            div.innerHTML = `
                <div style="font-weight:700;font-size:15px;margin-bottom:3px;">Yoğunluk Çemberi</div>
                <div style="font-size:11px;color:#666;margin-bottom:8px;">${countLabel}</div>
                <div style="display:flex;justify-content:space-between;gap:18px;font-size:14px;margin:5px 0;">
                    <span>Yarıçap</span><strong id="ayca-radius-value">1000 m</strong>
                </div>
                <div style="display:flex;justify-content:space-between;gap:18px;font-size:14px;margin:5px 0;">
                    <span>Çember içi</span><strong id="ayca-count-value">0 eczane</strong>
                </div>
                <div style="display:flex;justify-content:space-between;gap:18px;font-size:14px;margin:5px 0;">
                    <span>Toplam oran</span><strong id="ayca-share-value">0%</strong>
                </div>
                <div style="margin-top:9px;font-size:12px;font-weight:600;">Yarıçapı değiştir</div>
                <input id="ayca-radius-slider" type="range" min="100" max="2500" step="50" value="1000"
                    style="width:100%;margin-top:5px;accent-color:#C62828;">
                <div style="display:flex;justify-content:space-between;font-size:10px;color:#666;">
                    <span>100 m</span><span>2500 m</span>
                </div>
                <button id="ayca-center-mode" type="button"
                    style="width:100%;margin-top:9px;padding:7px 8px;border:1px solid #bbb;border-radius:7px;background:white;cursor:pointer;font-weight:600;">
                    Haritadan merkez seç
                </button>
                <div id="ayca-density-hint" style="margin-top:7px;font-size:11px;color:#666;line-height:1.3;">
                    Kırmızı noktayı sürükleyerek çemberi taşıyabilirsiniz.
                </div>`;

            div.style.background = 'rgba(255,255,255,.97)';
            div.style.border = '1px solid #d8d8d8';
            div.style.borderRadius = '10px';
            div.style.padding = '12px 14px';
            div.style.minWidth = '225px';
            div.style.boxShadow = '0 2px 8px rgba(0,0,0,.15)';
            div.style.fontFamily = 'Arial, sans-serif';
            div.style.color = '#222';

            L.DomEvent.disableClickPropagation(div);
            L.DomEvent.disableScrollPropagation(div);
            return div;
        }
    });

    map.addControl(new DensityControl());

    function distanceMeters(lat1, lon1, lat2, lon2) {
        const R = 6371000;
        const toRad = d => d * Math.PI / 180;
        const dLat = toRad(lat2 - lat1);
        const dLon = toRad(lon2 - lon1);
        const a = Math.sin(dLat / 2) ** 2 +
                  Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) *
                  Math.sin(dLon / 2) ** 2;
        return 2 * R * Math.asin(Math.sqrt(a));
    }

    function updateDensity() {
        const center = densityCircle.getLatLng();
        const radius = densityCircle.getRadius();
        let count = 0;

        pharmacies.forEach(p => {
            if (distanceMeters(center.lat, center.lng, p.lat, p.lon) <= radius) {
                count += 1;
            }
        });

        const radiusEl = document.getElementById('ayca-radius-value');
        const countEl = document.getElementById('ayca-count-value');
        const shareEl = document.getElementById('ayca-share-value');

        if (radiusEl) radiusEl.textContent = Math.round(radius) + ' m';
        if (countEl) countEl.textContent = count + ' eczane';
        if (shareEl) shareEl.textContent = total ? ((count / total) * 100).toFixed(1).replace('.', ',') + '%' : '0%';
    }

    centerHandle.on('drag', function(e) {
        densityCircle.setLatLng(e.target.getLatLng());
        updateDensity();
    });

    centerHandle.on('dragend', function(e) {
        densityCircle.setLatLng(e.target.getLatLng());
        updateDensity();
    });

    setTimeout(function() {
        const slider = document.getElementById('ayca-radius-slider');
        const centerButton = document.getElementById('ayca-center-mode');
        const hint = document.getElementById('ayca-density-hint');
        let chooseCenter = false;

        if (slider) {
            slider.addEventListener('input', function() {
                densityCircle.setRadius(Number(this.value));
                updateDensity();
            });
        }

        if (centerButton) {
            centerButton.addEventListener('click', function() {
                chooseCenter = !chooseCenter;
                if (chooseCenter) {
                    this.textContent = 'Haritada bir noktaya tıkla';
                    this.style.background = '#FDECEC';
                    this.style.borderColor = '#C62828';
                    if (hint) hint.textContent = 'Şimdi haritada çemberin merkezini istediğiniz yere tıklayın.';
                } else {
                    this.textContent = 'Haritadan merkez seç';
                    this.style.background = 'white';
                    this.style.borderColor = '#bbb';
                    if (hint) hint.textContent = 'Kırmızı noktayı sürükleyerek çemberi taşıyabilirsiniz.';
                }
            });
        }

        map.on('click', function(e) {
            if (!chooseCenter) return;

            centerHandle.setLatLng(e.latlng);
            densityCircle.setLatLng(e.latlng);
            updateDensity();

            chooseCenter = false;
            if (centerButton) {
                centerButton.textContent = 'Haritadan merkez seç';
                centerButton.style.background = 'white';
                centerButton.style.borderColor = '#bbb';
            }
            if (hint) hint.textContent = 'Merkez değiştirildi. Kırmızı noktayı da sürükleyebilirsiniz.';
        });

        updateDensity();
    }, 0);
})();
{% endmacro %}
"""
        )

        self.pharmacies_json = pharmacies_json
        self.count_label_json = count_label_json
        self.center_lat = center_lat
        self.center_lon = center_lon


def add_density_circle_widget(
    map_obj: folium.Map,
    count_df: pd.DataFrame,
    center_df: pd.DataFrame,
    count_label: str,
) -> None:
    """Haritaya çalışan yoğunluk çemberi kontrolünü ekler."""
    pharmacy_points = [
        {
            "name": str(row["Eczane"]),
            "lat": float(row["Latitude"]),
            "lon": float(row["Longitude"]),
        }
        for _, row in count_df.iterrows()
    ]

    control = DensityCircleControl(
        pharmacy_points=pharmacy_points,
        center_lat=float(center_df["Latitude"].median()),
        center_lon=float(center_df["Longitude"].median()),
        count_label=count_label,
    )
    control.add_to(map_obj)


def build_map(
    full_df: pd.DataFrame,
    selected_groups: set[str],
    count_df: pd.DataFrame,
    count_label: str,
) -> folium.Map:
    center = [
        float(full_df["Latitude"].median()),
        float(full_df["Longitude"].median()),
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

    selected_df = full_df[full_df["Grup"].isin(selected_groups)].copy()

    add_group_boundaries(map_obj, full_df, selected_groups)
    add_pharmacy_markers(map_obj, selected_df, selected_groups)
    add_density_circle_widget(map_obj, count_df, full_df, count_label)

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
                float(full_df["Latitude"].min()),
                float(full_df["Longitude"].min()),
            ],
            [
                float(full_df["Latitude"].max()),
                float(full_df["Longitude"].max()),
            ],
        ],
        padding=(25, 25),
    )

    return map_obj


# =========================================================
# SIDEBAR / ALT GRUP FİLTRELERİ
# =========================================================
ALL_GROUPS = [
    "A1", "A2", "A3", "A4",
    "B1", "B2", "B3", "B4",
    "C1", "C2", "C3", "C4",
    "D1", "D2", "D3", "D4",
]


def init_filter_state() -> None:
    for group_name in ALL_GROUPS:
        key = f"filter_{group_name}"
        if key not in st.session_state:
            st.session_state[key] = True


def select_all_groups() -> None:
    for group_name in ALL_GROUPS:
        st.session_state[f"filter_{group_name}"] = True


def clear_all_groups() -> None:
    for group_name in ALL_GROUPS:
        st.session_state[f"filter_{group_name}"] = False


def select_main_group(letter: str) -> None:
    for group_name in ALL_GROUPS:
        st.session_state[f"filter_{group_name}"] = group_name.startswith(letter)


def render_sidebar_filters(pharmacies: pd.DataFrame) -> tuple[set[str], str]:
    init_filter_state()

    st.sidebar.header("Alt Grup Filtreleri")
    st.sidebar.caption("Seçim değiştiğinde marker ve grup sınırı birlikte güncellenir.")

    c1, c2 = st.sidebar.columns(2)
    c1.button("Tümünü Aç", on_click=select_all_groups, use_container_width=True)
    c2.button("Temizle", on_click=clear_all_groups, use_container_width=True)

    st.sidebar.caption("Hızlı seçim")
    q1, q2, q3, q4 = st.sidebar.columns(4)
    q1.button("A", on_click=select_main_group, args=("A",), use_container_width=True)
    q2.button("B", on_click=select_main_group, args=("B",), use_container_width=True)
    q3.button("C", on_click=select_main_group, args=("C",), use_container_width=True)
    q4.button("D", on_click=select_main_group, args=("D",), use_container_width=True)

    selected: set[str] = set()

    for letter in ["A", "B", "C", "D"]:
        st.sidebar.markdown(f"**Grup {letter}**")
        cols = st.sidebar.columns(2)
        letter_groups = [f"{letter}{i}" for i in range(1, 5)]
        for i, group_name in enumerate(letter_groups):
            count = int((pharmacies["Grup"] == group_name).sum())
            checked = cols[i % 2].checkbox(
                f"{group_name} ({count})",
                key=f"filter_{group_name}",
            )
            if checked:
                selected.add(group_name)

    st.sidebar.divider()
    st.sidebar.subheader("Yoğunluk Çemberi")
    count_mode = st.sidebar.radio(
        "Çember hangi eczaneleri saysın?",
        ["Seçili alt gruplar", "Tüm eczaneler"],
        index=0,
    )

    selected_count = int(pharmacies["Grup"].isin(selected).sum())
    st.sidebar.metric("Haritada seçili eczane", selected_count)
    st.sidebar.caption(f"Seçili alt grup: {len(selected)} / {len(ALL_GROUPS)}")

    with st.sidebar.expander("Seçili alt gruplardaki eczaneler"):
        for group_name in ALL_GROUPS:
            if group_name not in selected:
                continue
            names = pharmacies.loc[
                pharmacies["Grup"] == group_name,
                "Eczane",
            ].astype(str).sort_values().tolist()
            st.markdown(
                f"**{group_name} ({len(names)}):** " + (", ".join(names) if names else "-")
            )

    return selected, count_mode


# =========================================================
# UYGULAMA
# =========================================================
st.title("Sivas Eczane Grup Haritası")
st.caption(
    "V3.0 — A1-D4 alt grupları sol panelden filtrelenebilir; marker ve sınırlar birlikte güncellenir. "
    "Yoğunluk çemberi seçili grupları veya tüm eczaneleri canlı olarak sayabilir."
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

    selected_groups, count_mode = render_sidebar_filters(pharmacies)
    selected_df = pharmacies[pharmacies["Grup"].isin(selected_groups)].copy()

    if count_mode == "Tüm eczaneler":
        count_df = pharmacies.copy()
        count_label = f"Tüm eczaneler sayılıyor ({len(count_df)})"
    else:
        count_df = selected_df.copy()
        count_label = f"Seçili alt gruplar sayılıyor ({len(count_df)})"

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Toplam eczane", len(pharmacies))
    col2.metric("Grubu eşleşen", assigned_count)
    col3.metric("Haritada seçili", len(selected_df))
    col4.metric("Grupsuz", missing_count)

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

    if not selected_groups:
        st.warning(
            "Hiç alt grup seçili değil. Harita tabanı ve yoğunluk çemberi görünür; "
            "eczane markerları ve grup sınırları gösterilmez."
        )

    pharmacy_map = build_map(
        full_df=pharmacies,
        selected_groups=selected_groups,
        count_df=count_df,
        count_label=count_label,
    )
    components.html(
        pharmacy_map.get_root().render(),
        height=900,
        scrolling=False,
    )

except Exception as exc:
    st.exception(exc)
