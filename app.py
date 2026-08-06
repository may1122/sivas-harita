import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import folium
from folium.plugins import MarkerCluster

st.set_page_config(
    page_title="Sivas Eczane Haritası",
    page_icon="🏥",
    layout="wide"
)

st.title("🏥 Sivas Eczane Grup Haritası")

# -----------------------------
# DOSYA İSİMLERİ
# -----------------------------
ECZANE_FILE = "nöbet-merkez tutan eczaneler(20260806-122152).xlsx"
GROUP_FILE = "SİVAS-GRUP(20260806-122152).xlsx"

# -----------------------------
# VERİYİ OKU
# -----------------------------
eczaneler = pd.read_excel(ECZANE_FILE)
gruplar = pd.read_excel(GROUP_FILE, header=None)

# -----------------------------
# Grup sözlüğü oluştur
# -----------------------------
grup_dict = {}

for col in gruplar.columns:
    grup = str(gruplar.iloc[0, col]).strip()

    for i in range(1, len(gruplar)):
        isim = gruplar.iloc[i, col]

        if pd.notna(isim):
            grup_dict[str(isim).strip().upper()] = grup

# -----------------------------
# Grup sütunu ekle
# -----------------------------
isim_kolon = eczaneler.columns[0]

eczaneler["GRUP"] = (
    eczaneler[isim_kolon]
    .astype(str)
    .str.upper()
    .map(grup_dict)
)

# -----------------------------
# Koordinat sütunlarını bul
# -----------------------------
lat_col = None
lon_col = None

for c in eczaneler.columns:
    cc = c.lower()

    if "lat" in cc:
        lat_col = c

    if "lon" in cc or "lng" in cc:
        lon_col = c

# -----------------------------
# Harita
# -----------------------------
m = folium.Map(
    location=[39.75, 37.02],
    zoom_start=13,
    tiles="CartoDB Positron"
)

cluster = MarkerCluster().add_to(m)

for _, row in eczaneler.iterrows():

    if pd.isna(row[lat_col]) or pd.isna(row[lon_col]):
        continue

    isim = row[isim_kolon]
    grup = row["GRUP"]

    folium.CircleMarker(
        location=[row[lat_col], row[lon_col]],
        radius=6,
        color="#1976D2",
        fill=True,
        fill_opacity=0.9,
        tooltip=f"<b>{isim}</b><br>{grup}"
    ).add_to(cluster)

# -----------------------------
# HTML göster
# -----------------------------
components.html(
    m._repr_html_(),
    height=850,
    scrolling=False
)
