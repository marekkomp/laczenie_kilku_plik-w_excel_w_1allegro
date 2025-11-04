# streamlit_merge_allegro_xlsm.py
# -*- coding: utf-8 -*-
"""
Łączenie wielu plików XLSM/XLSX z Allegro (arkusz "Szablon") w jeden wynik:
- Nagłówki w WIERSZU 4 (index 3). Wiersze 1–3 ignorowane.
- Kolejność kolumn = dokładnie kolejność z PIERWSZEGO pliku.
- Jeśli kolejne pliki mają dodatkowe kolumny, są DOPISYWANE na końcu.
- Brakujące kolumny w danym pliku są dodawane z pustą wartością.
- Zero „kanonicznych” nazw, zero fuzzy/deduplikacji – czyste scalenie po NAZWACH nagłówków.
- Eksport do CSV i XLSX.

Uwaga: Pandas wczytuje wszystko jako tekst (dtype=str), żeby nie psuć wartości.
"""

import io
import pandas as pd
import streamlit as st

SHEET_NAME = "Szablon"
HEADER_ROW_IDX = 3  # wiersz nagłówków = 4
FORCE_ALL_STR = True

st.set_page_config(page_title="Scal XLSM Allegro — Szablon", layout="wide")
st.title("Scal XLSM/XLSX (Allegro) — arkusz 'Szablon', nagłówki od wiersza 4")

files = st.file_uploader(
    "Wybierz 2–50 plików XLSM/XLSX (arkusz 'Szablon')",
    type=["xlsm", "xlsx", "xls"],
    accept_multiple_files=True,
)

if not files or len(files) < 2:
    st.info("Wgraj co najmniej 2 pliki.")
    st.stop()

frames = []
column_order = []  # kolejność końcowa wg pierwszego pliku + nowe kolumny na końcu

for idx, f in enumerate(files):
    try:
        if FORCE_ALL_STR:
            df = pd.read_excel(f, sheet_name=SHEET_NAME, header=HEADER_ROW_IDX, dtype=str)
        else:
            df = pd.read_excel(f, sheet_name=SHEET_NAME, header=HEADER_ROW_IDX)
    except Exception as e:
        st.error(f"Nie udało się wczytać pliku {f.name}: {e}")
        continue

    # Zamień całe NaN na puste stringi (spójny typ tekstowy)
    df = df.fillna("")

    # Ustal kolejność kolumn: najpierw z pierwszego pliku, potem dokładamy nowe
    if idx == 0:
        column_order = list(df.columns)
    else:
        for c in df.columns:
            if c not in column_order:
                column_order.append(c)

    frames.append(df)

if not frames:
    st.stop()

# Zrób unijny zestaw kolumn, ale ustaw docelową kolejność = column_order
all_columns = column_order.copy()

# Upewnij się, że każdy df ma wszystkie kolumny (dodaj brakujące)
prepared = []
for df in frames:
    missing = [c for c in all_columns if c not in df.columns]
    for m in missing:
        df[m] = ""
    # Przestaw w kolejności końcowej
    df = df[all_columns]
    prepared.append(df)

merged = pd.concat(prepared, axis=0, ignore_index=True)

st.subheader("Podgląd (pierwsze 200 wierszy)")
st.dataframe(merged.head(200), use_container_width=True)

st.markdown("### Eksport")
# CSV
csv_bytes = merged.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    "Pobierz CSV",
    data=csv_bytes,
    file_name="allegro_merged.csv",
    mime="text/csv",
)

# XLSX (z jedną zakładką)
output = io.BytesIO()
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    merged.to_excel(writer, index=False, sheet_name="merged")

st.download_button(
    "Pobierz XLSX",
    data=output.getvalue(),
    file_name="allegro_merged.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
