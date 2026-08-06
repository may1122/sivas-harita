#V1.1 - grup çakışması / kayma düzeltmesi
from __future__ import annotations

import html
import math
import re
import unicodedata
from collections import defaultdict
from pathlib import Path

import folium
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from folium import FeatureGroup
from folium.plugins import Fullscreen, MeasureControl
from shapely.geometry import MultiPoint, Point

# =========================================================
# SAYFA AYARLARI
# =========================================================
st.set_page_config(
    page_title="Sivas Eczane Grup Haritası",
    page_icon="💊",
    layout="wide",
)

BASE_DIR = Path(__file__).resolve().parent
ECZANE_FILE_NAME = "nöbet-merkez tutan eczaneler(20260806-122152).xlsx"
GROUP_FILE_NAME = "SİVAS-GRUP(20260806-122152).xlsx"

MAIN_COLORS = {
    "A": "#E53935",  # kırmızı
    "B": "#1E88E5",  # mavi
    "C": "#43A047",  # yeşil
    "D": "#FB8C00",  # turuncu
}

SUBGROUP_DASH = {
    "1": "4,5",
    "2": "8,5",
    "3": "12,5",
    "4": "2,5",
}

# Bir alt grup içinde, geri kalan noktalardan bu kat kadar uzak kalan
# tek nokta(lar), sınır çiziminden (convex hull) hariç tutulur.
# (Harita üzerindeki nokta/tooltip'ten hariç tutulmaz, sadece kesikli
# sınır çizgisini o kadar germesin diye.)
OUTLIER_DISTANCE_FACTOR = 3.0
OUTLIER_MIN_ABS_METERS = 600.0


def normalize_name(value: object) -> str:
    """Eczane adlarını Türkçe karakterleri koruyarak karşılaştırılabilir hale getirir."""
    if pd.isna(value):
        return ""
    text = unicodedata.normalize("NFKC", str(value)).strip().upper()
    text = re.sub(r"\s+", "", text)
    text = re.sub(r"[^0-9A-ZÇĞİÖŞÜ]", "", text)
    return text


def locate_file(exact_name: str, keyword: str) -> Path | None:
    """Önce tam dosya adını, sonra repo içindeki uygun xlsx dosyasını bulur."""
    exact = BASE_DIR / exact_name
    if exact.exists():
        return exact

    candidates = sorted(BASE_DIR.glob("*.xlsx"))
    keyword_norm = normalize_name(keyword)
    for candidate in candidates:
        if keyword_norm in normalize_name(candidate.name):
            return candidate
    return None


def parse_coordinate(value: object) -> tuple[float, float] | None:
    if pd.isna(value):
        return None
    text = str(value).strip().replace(";", ",")
    match = re.search(r"(-?\d{1,3}[\.,]\d+)\s*,\s*(-?\d{1,3}[\.,]\d+)", text)
    if not match:
        return None
    try:
        lat = float(match.group(1).replace(",", "."))
        lon = float(match.group(2).replace(",", "."))
    except ValueError:
        return None
    if not (-90 <= lat <= 90 and -180 <= lon <= 180):
        return None
    return lat, lon


@st.cache_data(show_spinner=False)
def read_pharmacies(path: str) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, engine="openpyxl")
    records: list[dict[str, object]] = []

    for _, row in raw.iterrows():
        if len(row) < 2:
            continue
        name = str(row.iloc[0]).strip() if not pd.isna(row.iloc[0]) else ""
        coord = parse_coordinate(row.iloc[1])
        if not name or normalize_name(name) in {"ECZANE", "ECZANEADI", "AD"} or coord is None:
            continue
        records.append(
            {
                "Eczane": name.strip(),
                "Anahtar": normalize_name(name),
                "Latitude": coord[0],
                "Longitude": coord[1],
            }
        )

    df = pd.DataFrame(records).drop_duplicates(subset=["Anahtar"], keep="first")
    if df.empty:
        raise ValueError("Koordinat dosyasında geçerli eczane ve koordinat bulunamadı.")
    return df


@st.cache_data(show_spinner=False)
def read_group_candidates(path: str) -> dict[str, list[str]]:
    """
    SİVAS-GRUP dosyasındaki A/C ve B/D bloklarını okur ve her isim için
    o isme rastlanan TÜM grup adaylarını (birden fazla olabilir) döndürür.
    Aynı isim birden fazla grupta geçiyorsa, çakışma coğrafi olarak
    resolve_group_conflicts() içinde çözülür.
    """
    raw = pd.read_excel(path, header=None, engine="openpyxl")
    candidates: dict[str, list[str]] = defaultdict(list)

    header_rows: list[tuple[int, list[tuple[int, str]]]] = []
    for row_idx, row in raw.iterrows():
        headers: list[tuple[int, str]] = []
        for col_idx, value in enumerate(row.tolist()):
            text = str(value).strip().upper() if not pd.isna(value) else ""
            if re.fullmatch(r"[ABCD][1-4]", text):
                headers.append((col_idx, text))
        if headers:
            header_rows.append((row_idx, headers))

    for block_index, (header_row_idx, headers) in enumerate(header_rows):
        next_header_row = (
            header_rows[block_index + 1][0]
            if block_index + 1 < len(header_rows)
            else len(raw)
        )

        headers = sorted(headers, key=lambda item: item[0])
        for header_pos, (start_col, group_name) in enumerate(headers):
            next_col = headers[header_pos + 1][0] if header_pos + 1 < len(headers) else len(raw.columns)
            # Aradaki geniş boşluğu karşı blok sanmamak için en fazla iki sütun oku.
            end_col = min(next_col, start_col + 2)

            for row_idx in range(header_row_idx + 1, next_header_row):
                row_values = raw.iloc[row_idx]
                # Yeni A/B/C/D ana başlığına gelindiyse blok sona ermiştir.
                first_cells = [str(v).strip().upper() for v in row_values.tolist() if not pd.isna(v)]
                if any(re.fullmatch(r"[ABCD]", value) for value in first_cells):
                    break

                for col_idx in range(start_col, end_col):
                    if col_idx >= len(raw.columns):
                        continue
                    value = raw.iat[row_idx, col_idx]
                    key = normalize_name(value)
                    if not key or re.fullmatch(r"[ABCD][1-4]?", key):
                        continue
                    candidates[key].append(group_name)

    if not candidates:
        raise ValueError("Grup dosyasında A1–D4 grup yapısı bulunamadı.")
    return dict(candidates)


def resolve_group_conflicts(
    candidates: dict[str, list[str]], pharmacies: pd.DataFrame
) -> tuple[dict[str, str], list[dict[str, str]]]:
    """
    Bir eczane adı Excel'de birden fazla farklı grupta geçiyorsa (kopyala-yapıştır
    hatası vb.), o eczanenin GERÇEK koordinatına en yakın grup merkezine göre
    çakışmayı çözer. Bu, tek bir yanlış-gruplanmış eczanenin sınır çizgisini
    haritanın öbür ucuna germesini (kaymasını) engeller.
    """
    coord_lookup = pharmacies.set_index("Anahtar")[["Latitude", "Longitude"]]

    # 1) Belirsiz olmayan (tek gruba ait) isimlerden her grubun kaba merkezini kur.
    unambiguous = {key: opts[0] for key, opts in candidates.items() if len(set(opts)) == 1}

    group_points: dict[str, list[tuple[float, float]]] = defaultdict(list)
    for key, group in unambiguous.items():
        if key in coord_lookup.index:
            lat, lon = coord_lookup.loc[key, ["Latitude", "Longitude"]]
            group_points[group].append((float(lat), float(lon)))

    centroids = {
        group: (
            sum(p[0] for p in pts) / len(pts),
            sum(p[1] for p in pts) / len(pts),
        )
        for group, pts in group_points.items()
        if pts
    }

    # 2) Çakışmaları koordinata en yakın merkeze göre çöz.
    final_map: dict[str, str] = dict(unambiguous)
    conflicts_resolved: list[dict[str, str]] = []

    for key, opts in candidates.items():
        unique_groups = sorted(set(opts))
        if len(unique_groups) <= 1:
            continue

        if key not in coord_lookup.index:
            # Koordinatı olmayan bir eczane için ilk görülen grubu kullan.
            final_map[key] = opts[0]
            continue

        lat, lon = coord_lookup.loc[key, ["Latitude", "Longitude"]]
        lat, lon = float(lat), float(lon)

        scored = [
            (g, (lat - centroids[g][0]) ** 2 + (lon - centroids[g][1]) ** 2)
            for g in unique_groups
            if g in centroids
        ]
        best_group = min(scored, key=lambda item: item[1])[0] if scored else opts[0]

        final_map[key] = best_group
        conflicts_resolved.append(
            {
                "isim": key,
                "excelde_gecen_gruplar": ", ".join(unique_groups),
                "secilen_grup": best_group,
            }
        )

    return final_map, conflicts_resolved


def meters_to_lat(meters: float) -> float:
    return meters / 111_320.0


def meters_to_lon(meters: float, latitude: float) -> float:
    return meters / (111_320.0 * max(math.cos(math.radians(latitude)), 0.2))


def haversine_meters(p1: tuple[float, float], p2: tuple[float, float]) -> float:
    lat1, lon1 = p1
    lat2, lon2 = p2
    r = 6_371_000.0
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlambda / 2) ** 2
    return 2 * r * math.asin(min(1.0, math.sqrt(a)))


def destination(center: tuple[float, float], bearing_deg: float, distance_m: float) -> tuple[float, float]:
    """Harita kılavuz çizgileri için yeterli hassasiyette düzlem yaklaşımı."""
    lat, lon = center
    angle = math.radians(bearing_deg)
    dlat = meters_to_lat(distance_m * math.cos(angle))
    dlon = meters_to_lon(distance_m * math.sin(angle), lat)
    return lat + dlat, lon + dlon


def add_sector_guides(
    map_obj: folium.Map,
    center: tuple[float, float],
    max_radius_m: float,
) -> None:
    """Dört ana bölgeyi pasta dilimi ve iç içe çemberlerle gösterir."""
    guides = FeatureGroup(name="Ana bölge pasta/çember kılavuzu", show=True)

    # 4 iç içe çember
    for fraction in (0.25, 0.50, 0.75, 1.0):
        folium.Circle(
            location=center,
            radius=max_radius_m * fraction,
            color="#5F6368",
            weight=1.2,
            opacity=0.55,
            fill=False,
            dash_array="6,6",
            tooltip=f"Merkezden yaklaşık {max_radius_m * fraction / 1000:.1f} km",
        ).add_to(guides)

    # Kuzeydoğu=A, güneydoğu=C, güneybatı=D, kuzeybatı=B şeklinde dört dilim.
    sector_defs = [
        ("A", 0, 90),
        ("C", 90, 180),
        ("D", 180, 270),
        ("B", 270, 360),
    ]
    for main_group, start_angle, end_angle in sector_defs:
        points = [center]
        for angle in range(start_angle, end_angle + 1, 3):
            points.append(destination(center, angle, max_radius_m))
        points.append(center)
        folium.Polygon(
            locations=points,
            color=MAIN_COLORS[main_group],
            weight=2,
            opacity=0.55,
            fill=True,
            fill_color=MAIN_COLORS[main_group],
            fill_opacity=0.035,
            tooltip=f"{main_group} ana bölgesi",
        ).add_to(guides)

    # Dört ayırıcı çizgi
    for angle in (0, 90, 180, 270):
        folium.PolyLine(
            [center, destination(center, angle, max_radius_m)],
            color="#424242",
            weight=1.5,
            opacity=0.55,
            dash_array="7,7",
        ).add_to(guides)

    guides.add_to(map_obj)


def split_outliers(group_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Bir alt grubun noktalarını, sınır (convex hull) çiziminde kullanılacak
    "çekirdek" noktalar ve o grubun geri kalanından anormal uzak kalan
    "aykırı" noktalar olarak ikiye ayırır. Aykırı noktalar haritada işaretçi
    olarak görünmeye devam eder, sadece kesikli sınır çizgisini germez.
    """
    if len(group_df) < 4:
        return group_df, group_df.iloc[0:0]

    lat_c = group_df["Latitude"].median()
    lon_c = group_df["Longitude"].median()
    center = (lat_c, lon_c)

    distances = group_df.apply(
        lambda r: haversine_meters((r["Latitude"], r["Longitude"]), center), axis=1
    )
    median_dist = distances.median()
    threshold = max(median_dist * OUTLIER_DISTANCE_FACTOR, OUTLIER_MIN_ABS_METERS)

    is_outlier = distances > threshold
    # Grubun yarısından fazlası "aykırı" çıkarsa muhtemelen asıl dağınık olan
    # gruptur, hiçbir şeyi hariç tutma (güvenlik freni).
    if is_outlier.sum() >= len(group_df) / 2:
        return group_df, group_df.iloc[0:0]

    return group_df[~is_outlier], group_df[is_outlier]


def add_group_boundaries(map_obj: folium.Map, df: pd.DataFrame) -> list[dict[str, str]]:
    """Her alt grubun eczanelerini çevreleyen renkli sınırlar çizer.
    Aşırı uzak kalan (aykırı) noktaları hariç tutarak sınırın haritanın
    öbür ucuna kaymasını engeller. Hariç tutulanların listesini döndürür.
    """
    boundary_layer = FeatureGroup(name="Alt grup sınırları", show=True)
    excluded_points: list[dict[str, str]] = []

    for group_name, group_df in df.dropna(subset=["Grup"]).groupby("Grup"):
        core_df, outlier_df = split_outliers(group_df)
        for _, row in outlier_df.iterrows():
            excluded_points.append({"eczane": row["Eczane"], "grup": str(group_name)})

        coords = list(zip(core_df["Longitude"], core_df["Latitude"]))
        if not coords:
            continue

        main_group = str(group_name)[0]
        subgroup_number = str(group_name)[1]
        color = MAIN_COLORS.get(main_group, "#616161")

        geometry = MultiPoint(coords).convex_hull
        # Az noktalı veya çizgisel gruplara görünür alan kazandır.
        if geometry.geom_type in {"Point", "LineString"}:
            geometry = geometry.buffer(0.0032)
        else:
            geometry = geometry.buffer(0.0015)
        geometry = geometry.simplify(0.00025)

        polygons = [geometry] if geometry.geom_type == "Polygon" else list(geometry.geoms)
        for polygon in polygons:
            lat_lon = [(lat, lon) for lon, lat in polygon.exterior.coords]
            folium.Polygon(
                locations=lat_lon,
                color=color,
                weight=2.5,
                opacity=0.9,
                dash_array=SUBGROUP_DASH.get(subgroup_number, "5,5"),
                fill=True,
                fill_color=color,
                fill_opacity=0.07,
                tooltip=f"{group_name} sınırı",
            ).add_to(boundary_layer)

        center = geometry.centroid
        folium.Marker(
            location=[center.y, center.x],
            icon=folium.DivIcon(
                html=(
                    f'<div style="font-size:13px;font-weight:800;color:{color};'
                    'text-shadow:0 0 3px white,0 0 3px white;white-space:nowrap;">'
                    f'{html.escape(str(group_name))}</div>'
                )
            ),
        ).add_to(boundary_layer)

    boundary_layer.add_to(map_obj)
    return excluded_points


def build_map(df: pd.DataFrame) -> tuple[folium.Map, list[dict[str, str]]]:
    center = (float(df["Latitude"].median()), float(df["Longitude"].median()))

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

    # En uzak eczaneye göre pasta/çember yarıçapı.
    center_point = Point(center[1], center[0])
    max_degree = max(
        center_point.distance(Point(lon, lat))
        for lat, lon in zip(df["Latitude"], df["Longitude"])
    )
    max_radius_m = max(2500.0, max_degree * 111_320.0 * 1.08)

    add_sector_guides(map_obj, center, max_radius_m)
    excluded_points = add_group_boundaries(map_obj, df)

    pharmacy_layer = FeatureGroup(name="Eczaneler", show=True)
    for _, row in df.iterrows():
        group = row.get("Grup")
        group_text = str(group) if pd.notna(group) else "Grupsuz"
        main_group = group_text[0] if group_text and group_text[0] in MAIN_COLORS else ""
        color = MAIN_COLORS.get(main_group, "#757575")

        tooltip = (
            '<div style="font-size:14px;line-height:1.35;">'
            f'<b>{html.escape(str(row["Eczane"]))}</b><br>'
            f'Grup: <b>{html.escape(group_text)}</b>'
            '</div>'
        )

        folium.CircleMarker(
            location=[row["Latitude"], row["Longitude"]],
            radius=5.8,
            color="#FFFFFF",
            weight=1.5,
            fill=True,
            fill_color=color,
            fill_opacity=0.95,
            tooltip=folium.Tooltip(tooltip, sticky=True, direction="top"),
        ).add_to(pharmacy_layer)

    pharmacy_layer.add_to(map_obj)
    Fullscreen(position="topright", title="Tam ekran", title_cancel="Tam ekrandan çık").add_to(map_obj)
    MeasureControl(position="topright", primary_length_unit="meters").add_to(map_obj)
    folium.LayerControl(collapsed=False, position="topright").add_to(map_obj)

    # Tüm eczaneleri ekrana sığdır.
    map_obj.fit_bounds(
        [
            [float(df["Latitude"].min()), float(df["Longitude"].min())],
            [float(df["Latitude"].max()), float(df["Longitude"].max())],
        ],
        padding=(25, 25),
    )
    return map_obj, excluded_points


# =========================================================
# UYGULAMA
# =========================================================
st.title("Sivas Eczane Grup Haritası")
st.caption(
    "Eczane isimleri sürekli görünmez. Noktanın üzerine gelince eczane adı ve grubu açılır."
)

pharmacy_path = locate_file(ECZANE_FILE_NAME, "nöbet-merkez tutan eczaneler")
group_path = locate_file(GROUP_FILE_NAME, "SİVAS-GRUP")

if pharmacy_path is None or group_path is None:
    st.error(
        "Excel dosyaları GitHub reposunda bulunamadı. app.py ile aynı klasöre aşağıdaki "
        "iki dosyayı eksiksiz yükleyin:"
    )
    st.code(f"{ECZANE_FILE_NAME}\n{GROUP_FILE_NAME}")
    st.stop()

try:
    pharmacies = read_pharmacies(str(pharmacy_path))
    group_candidates = read_group_candidates(str(group_path))
    group_map, conflicts_resolved = resolve_group_conflicts(group_candidates, pharmacies)
    pharmacies["Grup"] = pharmacies["Anahtar"].map(group_map)

    # Görsel kontrolde kolaylık için grubu olmayanları ayrıca göster.
    missing_count = int(pharmacies["Grup"].isna().sum())
    assigned_count = len(pharmacies) - missing_count

    col1, col2, col3 = st.columns(3)
    col1.metric("Toplam eczane", len(pharmacies))
    col2.metric("Grubu eşleşen", assigned_count)
    col3.metric("Grupsuz", missing_count)

    if missing_count:
        missing_names = ", ".join(pharmacies.loc[pharmacies["Grup"].isna(), "Eczane"].tolist())
        with st.expander(f"Grubu bulunamayan {missing_count} eczaneyi göster"):
            st.write(missing_names)

    if conflicts_resolved:
        with st.expander(
            f"⚠️ Grup dosyasında birden fazla grupta geçen {len(conflicts_resolved)} eczane bulundu "
            "(en yakın konuma göre otomatik çözüldü)"
        ):
            st.dataframe(pd.DataFrame(conflicts_resolved), hide_index=True, use_container_width=True)
            st.caption(
                "Bu isimler SİVAS-GRUP dosyasında birden fazla grup altında yazılmış. "
                "Eczanenin gerçek koordinatına en yakın grup merkezi otomatik seçildi. "
                "Doğru olduğundan emin olmak için Excel dosyasındaki mükerrer kayıtları kontrol edin."
            )

    pharmacy_map, excluded_points = build_map(pharmacies)

    if excluded_points:
        with st.expander(
            f"ℹ️ {len(excluded_points)} eczane, kendi grubunun geri kalanından çok uzak "
            "olduğu için sınır çiziminden hariç tutuldu"
        ):
            st.dataframe(pd.DataFrame(excluded_points), hide_index=True, use_container_width=True)
            st.caption(
                "Bu eczaneler haritada nokta olarak görünmeye devam ediyor, sadece kesikli "
                "sınır çizgisini o kadar uzağa germiyor. Koordinat veya grup ataması hatalı "
                "olabilir, kontrol etmenizi öneririz."
            )

    components.html(pharmacy_map.get_root().render(), height=900, scrolling=False)

except Exception as exc:
    st.exception(exc)
