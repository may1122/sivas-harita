from __future__ import annotations

import html
import re
import unicodedata
from pathlib import Path

import folium
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from folium import FeatureGroup
from folium.plugins import Fullscreen, MeasureControl

# =========================================================
# AYÇA UŞAK ECZANE HARİTASI
# VERSION : V1.4 - ALT GRUP BÖLGE ÇİZGİLERİ
# DATE    : 12.08.2026
# =========================================================

st.set_page_config(
    page_title="Uşak Eczane Haritası",
    page_icon="💊",
    layout="wide",
)

BASE_DIR = Path(__file__).resolve().parent
PREFERRED_FILE_NAME = "nöbet_ecz_liste_koordinatli(2).xlsx"

# ---------------------------------------------------------
# GRUPLAR
# A = YEŞİL, B = MAVİ, C = KIRMIZI
# ---------------------------------------------------------

GROUPS: dict[str, list[str]] = {
    "A1": [
        "AKKAYA", "HAZAL", "ERDEM", "ACAR", "ALTINPINAR", "LOKMAN", "HİLAL", "AKTAY", "YAŞAM", "ÇAKIR",
    ],
    "A2": [
        "SU", "ZAFER", "NUR", "AHSEN", "IŞIL", "DİDEM", "SAĞLIK",
        "TAN", "FİLİZ", "YEŞİM",
    ],
    "A3": [
        "ZÜMRÜT", "AYKANAT", "AKŞAHİN", "IHLAMUR", "AKINCI",
        "ÖMÜR", "DÖNMEZ", "ÇAVUSOĞLU", "İREM", "AVGAN",
    ],

    "B1": [
        "FATİH", "YENİ ŞİFA", "AYAN", "EYMEN", "GÜRAN",
        "MUTAFOĞLU", "YÜKSEL", "ÖZLEM", "EGE", "DOKUR", "BALKAN", 
    ],
    "B2": [
        "ÖRNEK", "SEVİM", "MASAL DİYARI", "PERDAHCI", "SULTAN",
        "AKDAĞ", "YAĞIZ", "SEVİNÇ", "BATI", "YAVUZ", "ÖZSEZER",
    ],
    "B3": [
        "SAMANCI", "KİRAZ", "DEMET", "SERAP", "DOĞA",
        "VİTAMİN", "GÜLŞİFA", "ÇEKİÇ", "ŞEYMA", "GÜNEŞ", "MURAT",
    ],

    "C1": [
        "AYDOĞDU", "ŞAN", "MERT", "GÜVEN", "BAŞER", "İLKE",
        "FERAH", "DURAN", "UMUT", "MAYA", "SERKAN",
    ],
    "C2": [
        "YILDIZ", "FARUK", "EGE HAYAT", "DAMLA", "GÖKSEL",
        "DEMİR", "GÜL", "NUSRET", "SELCEN", "EBRU",
    ],
    "C3": [
        "DÜLGEROĞLU", "SÜMER", "AŞİYAN", "POYRAZ", "EYLÜL",
        "MENDEPAZARI", "ÖZÇELİK", "BİZİM", "ÖZYAVUZ", "HUZUR",
    ],
}


GROUP_COLORS = {
    "A": "#16A34A",  # yeşil
    "B": "#2563EB",  # mavi
    "C": "#DC2626",  # kırmızı
}

GROUP_LABELS = {
    "A": "Grup A - Yeşil",
    "B": "Grup B - Mavi",
    "C": "Grup C - Kırmızı",
}


def normalize_text(value: object) -> str:
    """Dosya, sütun ve eczane adlarını toleranslı eşleştirmek için normalize eder."""
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
    return re.sub(r"[^0-9A-Z]", "", text)


def build_group_lookup() -> dict[str, str]:
    """Her eczane anahtarını tek bir gruba bağlar ve çift atamayı engeller."""
    lookup: dict[str, str] = {}
    duplicates: list[str] = []

    for group_name, pharmacy_names in GROUPS.items():
        for pharmacy_name in pharmacy_names:
            key = normalize_text(pharmacy_name)
            if not key:
                continue

            if key in lookup and lookup[key] != group_name:
                duplicates.append(pharmacy_name)
                continue

            lookup[key] = group_name

    if duplicates:
        raise ValueError(
            "Birden fazla gruba atanmış eczane var: " + ", ".join(sorted(set(duplicates)))
        )

    return lookup


GROUP_LOOKUP = build_group_lookup()


def locate_pharmacy_file() -> Path | None:
    """Önce tercih edilen adı, sonra uygun koordinat Excel'ini otomatik bulur."""
    preferred = BASE_DIR / PREFERRED_FILE_NAME
    if preferred.exists():
        return preferred

    candidates = sorted(BASE_DIR.glob("*.xlsx"))

    for candidate in candidates:
        key = normalize_text(candidate.name)
        if "NOBET" in key and "KOORDINAT" in key:
            return candidate

    for candidate in candidates:
        key = normalize_text(candidate.name)
        if "ECZ" in key or "ECZANE" in key:
            return candidate

    return candidates[0] if candidates else None


def find_column(columns: list[object], aliases: list[str]) -> object | None:
    """Sütun adını farklı yazım ihtimallerine rağmen bulur."""
    normalized = {normalize_text(col): col for col in columns}

    for alias in aliases:
        alias_key = normalize_text(alias)
        if alias_key in normalized:
            return normalized[alias_key]

    for alias in aliases:
        alias_key = normalize_text(alias)
        for key, original in normalized.items():
            if alias_key and alias_key in key:
                return original

    return None


@st.cache_data(show_spinner=False)
def read_pharmacies(path: str, file_version: int) -> pd.DataFrame:
    """Uşak koordinat Excel'ini okuyup standart sütun yapısına çevirir."""
    del file_version

    excel = pd.ExcelFile(path, engine="openpyxl")

    selected_df: pd.DataFrame | None = None
    selected_sheet = ""

    for sheet_name in excel.sheet_names:
        candidate = pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")
        columns = list(candidate.columns)

        name_col = find_column(columns, ["Eczane Adı", "Eczane", "Ad"])
        lat_col = find_column(columns, ["Enlem (Lat)", "Enlem", "Latitude", "Lat"])
        lng_col = find_column(columns, ["Boylam (Lng)", "Boylam", "Longitude", "Lng", "Lon"])

        if name_col is not None and lat_col is not None and lng_col is not None:
            selected_df = candidate.copy()
            selected_sheet = sheet_name
            break

    if selected_df is None:
        raise ValueError(
            "Excel içinde Eczane Adı, Enlem ve Boylam sütunlarını içeren uygun sayfa bulunamadı."
        )

    columns = list(selected_df.columns)
    name_col = find_column(columns, ["Eczane Adı", "Eczane", "Ad"])
    phone_col = find_column(columns, ["Telefon", "Tel", "Phone"])
    address_col = find_column(columns, ["Adres", "Address"])
    lat_col = find_column(columns, ["Enlem (Lat)", "Enlem", "Latitude", "Lat"])
    lng_col = find_column(columns, ["Boylam (Lng)", "Boylam", "Longitude", "Lng", "Lon"])

    result = pd.DataFrame()
    result["Eczane"] = selected_df[name_col].astype(str).str.strip()
    result["Latitude"] = pd.to_numeric(selected_df[lat_col], errors="coerce")
    result["Longitude"] = pd.to_numeric(selected_df[lng_col], errors="coerce")

    if phone_col is not None:
        result["Telefon"] = selected_df[phone_col].fillna("").astype(str).str.strip()
    else:
        result["Telefon"] = ""

    if address_col is not None:
        result["Adres"] = selected_df[address_col].fillna("").astype(str).str.strip()
    else:
        result["Adres"] = ""

    result = result[
        result["Eczane"].ne("")
        & result["Latitude"].between(-90, 90)
        & result["Longitude"].between(-180, 180)
    ].copy()

    result["Anahtar"] = result["Eczane"].map(normalize_text)
    result = result.drop_duplicates(subset=["Anahtar"], keep="first").reset_index(drop=True)

    # Grup ataması
    result["Alt Grup"] = result["Anahtar"].map(GROUP_LOOKUP).fillna("ATANMAMIŞ")
    result["Grup"] = result["Alt Grup"].where(
        result["Alt Grup"].eq("ATANMAMIŞ"),
        result["Alt Grup"].str[0],
    )

    result.attrs["sheet_name"] = selected_sheet

    if result.empty:
        raise ValueError("Excel'de haritada gösterilebilecek geçerli eczane koordinatı bulunamadı.")

    return result


def add_pharmacy_markers(map_obj: folium.Map, df: pd.DataFrame) -> None:
    """Eczaneleri A/B/C gruplarına göre ayrı renk ve katmanlarda gösterir."""

    layers = {
        "A": FeatureGroup(name=GROUP_LABELS["A"], show=True),
        "B": FeatureGroup(name=GROUP_LABELS["B"], show=True),
        "C": FeatureGroup(name=GROUP_LABELS["C"], show=True),
        "ATANMAMIŞ": FeatureGroup(name="Atanmamış", show=True),
    }

    for _, row in df.iterrows():
        group = str(row["Grup"])
        subgroup = str(row.get("Alt Grup", "ATANMAMIŞ"))
        pharmacy_name = html.escape(str(row["Eczane"]))
        phone = html.escape(str(row.get("Telefon", "")))
        address = html.escape(str(row.get("Adres", "")))

        if group in GROUP_COLORS:
            marker_color = GROUP_COLORS[group]
            group_text = GROUP_LABELS[group]
        else:
            marker_color = "#6B7280"
            group_text = "Atanmamış"

        tooltip_html = (
            '<div style="font-size:14px; line-height:1.4;">'
            f"<b>{pharmacy_name}</b><br>"
            f'<span style="color:{marker_color}; font-weight:700;">{group_text} / {subgroup}</span>'
            "</div>"
        )

        popup_parts = [
            '<div style="font-family:Arial,sans-serif; font-size:13px; line-height:1.5; min-width:230px;">',
            f'<div style="font-size:15px; font-weight:700; margin-bottom:4px;">{pharmacy_name}</div>',
            f'<div style="font-weight:700; color:{marker_color}; margin-bottom:6px;">{group_text} / Alt Grup {subgroup}</div>',
        ]

        if address:
            popup_parts.append(f"<div><b>Adres:</b> {address}</div>")
        if phone:
            popup_parts.append(f"<div><b>Telefon:</b> {phone}</div>")

        popup_parts.append(
            f'<div style="margin-top:6px; color:#666;">'
            f'{float(row["Latitude"]):.6f}, {float(row["Longitude"]):.6f}'
            "</div>"
        )
        popup_parts.append("</div>")

        folium.CircleMarker(
            location=[float(row["Latitude"]), float(row["Longitude"])],
            radius=6.5,
            color="#FFFFFF",
            weight=1.7,
            fill=True,
            fill_color=marker_color,
            fill_opacity=0.97,
            tooltip=folium.Tooltip(
                tooltip_html,
                sticky=True,
                direction="top",
            ),
            popup=folium.Popup("".join(popup_parts), max_width=380),
        ).add_to(layers[group])

    for layer in layers.values():
        layer.add_to(map_obj)


def add_legend(map_obj: folium.Map) -> None:
    legend_html = """
    <div style="
        position: fixed;
        bottom: 30px;
        left: 30px;
        z-index: 9999;
        background: white;
        border: 1px solid #cfcfcf;
        border-radius: 8px;
        padding: 10px 14px;
        font-family: Arial, sans-serif;
        font-size: 13px;
        box-shadow: 0 1px 6px rgba(0,0,0,0.18);
    ">
        <div style="font-weight:700; margin-bottom:7px;">Uşak Eczane Grupları</div>
        <div><span style="display:inline-block;width:11px;height:11px;border-radius:50%;background:#16A34A;margin-right:7px;"></span>Grup A</div>
        <div><span style="display:inline-block;width:11px;height:11px;border-radius:50%;background:#2563EB;margin-right:7px;"></span>Grup B</div>
        <div><span style="display:inline-block;width:11px;height:11px;border-radius:50%;background:#DC2626;margin-right:7px;"></span>Grup C</div>
    </div>
    """
    map_obj.get_root().html.add_child(folium.Element(legend_html))



def convex_hull(points: list[tuple[float, float]]) -> list[tuple[float, float]]:
    """Noktaların dış sınırını (convex hull) hesaplar."""
    points = sorted(set(points))
    if len(points) <= 1:
        return points

    def cross(o, a, b):
        return (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0])

    lower = []
    for p in points:
        while len(lower) >= 2 and cross(lower[-2], lower[-1], p) <= 0:
            lower.pop()
        lower.append(p)

    upper = []
    for p in reversed(points):
        while len(upper) >= 2 and cross(upper[-2], upper[-1], p) <= 0:
            upper.pop()
        upper.append(p)

    return lower[:-1] + upper[:-1]


def add_subgroup_boundaries(map_obj: folium.Map, df: pd.DataFrame) -> None:
    """A1-A3, B1-B3, C1-C3 alt gruplarını çizgilerle sınırlar."""
    subgroup_names = ["A1", "A2", "A3", "B1", "B2", "B3", "C1", "C2", "C3"]

    for subgroup in subgroup_names:
        subset = df[df["Alt Grup"] == subgroup].copy()
        if subset.empty:
            continue

        main_group = subgroup[0]
        line_color = GROUP_COLORS.get(main_group, "#6B7280")

        layer = FeatureGroup(
            name=f"{subgroup} bölge sınırı",
            show=True,
        )

        points = [
            (float(row["Latitude"]), float(row["Longitude"]))
            for _, row in subset.iterrows()
        ]

        hull = convex_hull(points)

        if len(hull) >= 3:
            folium.Polygon(
                locations=hull,
                color=line_color,
                weight=3,
                opacity=0.85,
                fill=True,
                fill_color=line_color,
                fill_opacity=0.06,
                tooltip=f"{subgroup} bölgesi",
            ).add_to(layer)

        elif len(hull) == 2:
            folium.PolyLine(
                locations=hull,
                color=line_color,
                weight=3,
                opacity=0.85,
                tooltip=f"{subgroup} bağlantısı",
            ).add_to(layer)

        elif len(hull) == 1:
            folium.Circle(
                location=hull[0],
                radius=120,
                color=line_color,
                weight=3,
                opacity=0.85,
                fill=False,
                tooltip=f"{subgroup} bölgesi",
            ).add_to(layer)

        layer.add_to(map_obj)


def build_map(df: pd.DataFrame) -> folium.Map:
    center = [
        float(df["Latitude"].median()),
        float(df["Longitude"].median()),
    ]

    map_obj = folium.Map(
        location=center,
        zoom_start=14,
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

    add_subgroup_boundaries(map_obj, df)
    add_pharmacy_markers(map_obj, df)
    add_legend(map_obj)

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
            [float(df["Latitude"].min()), float(df["Longitude"].min())],
            [float(df["Latitude"].max()), float(df["Longitude"].max())],
        ],
        padding=(25, 25),
    )

    return map_obj


# =========================================================
# UYGULAMA
# =========================================================

st.title("Uşak Eczane Haritası")
st.caption(
    "AYÇA — Grup A yeşil, Grup B mavi, Grup C kırmızı; alt gruplar A1-A3, B1-B3, C1-C3. "
    "Aynı alt gruptaki eczaneler bölge sınır çizgileriyle çevrelenir. "
    "Eczane adları fareyle üzerine gelince görünür; tıklayınca detay açılır."
)

pharmacy_path = locate_pharmacy_file()

if pharmacy_path is None:
    st.error(
        "Koordinat Excel dosyası GitHub reposunda bulunamadı. "
        "app.py ile aynı klasöre Uşak koordinat Excel dosyasını yükleyin."
    )
    st.stop()

try:
    pharmacies = read_pharmacies(
        str(pharmacy_path),
        pharmacy_path.stat().st_mtime_ns,
    )

    group_counts = pharmacies["Grup"].value_counts()

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Toplam eczane", len(pharmacies))
    col2.metric("Grup A", int(group_counts.get("A", 0)))
    col3.metric("Grup B", int(group_counts.get("B", 0)))
    col4.metric("Grup C", int(group_counts.get("C", 0)))
    col5.metric("Atanmamış", int(group_counts.get("ATANMAMIŞ", 0)))

    unassigned = pharmacies[pharmacies["Grup"] == "ATANMAMIŞ"]
    if not unassigned.empty:
        st.warning(
            "Gruba atanmamış eczaneler: "
            + ", ".join(unassigned["Eczane"].astype(str).tolist())
        )

    with st.expander("Grup dağılımı"):
        for group_name in ["A", "B", "C"]:
            names = pharmacies.loc[
                pharmacies["Grup"] == group_name, "Eczane"
            ].sort_values().tolist()
            st.markdown(
                f"**Grup {group_name} ({len(names)} eczane):** "
                + ", ".join(names)
            )

    with st.expander("Alt grup dağılımı"):
        subgroup_counts = pharmacies["Alt Grup"].value_counts()
        for subgroup_name in ["A1", "A2", "A3", "B1", "B2", "B3", "C1", "C2", "C3"]:
            names = pharmacies.loc[
                pharmacies["Alt Grup"] == subgroup_name, "Eczane"
            ].sort_values().tolist()
            st.markdown(
                f"**{subgroup_name} ({len(names)} eczane):** "
                + ", ".join(names)
            )

    with st.expander("Kullanılan veri dosyası"):
        st.write(f"Dosya: `{pharmacy_path.name}`")
        st.write(f"Sayfa: `{pharmacies.attrs.get('sheet_name', '-')}`")

    pharmacy_map = build_map(pharmacies)
    components.html(
        pharmacy_map.get_root().render(),
        height=900,
        scrolling=False,
    )

except Exception as exc:
    st.exception(exc)
