import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title="PO PDF → Excel", page_icon="📦", layout="centered")

st.title("📦 PO Extractor – Vendor Style ID / Units / Cost")
st.markdown("Wgraj plik PDF z Purchase Order (Levi's lub Nike/Macy's), a aplikacja wyciągnie dane i pozwoli pobrać Excel.")

uploaded_file = st.file_uploader("Wybierz plik PDF", type="pdf")

# ── LEVI'S parser ──────────────────────────────────────────────────────────────
# Linia: 000003785191L359-C1E-NH VINT INDIG ... 6 EA 3.50 3.50 21.00
LEVIS_PATTERN = re.compile(
    r'^\d{10}'
    r'([A-Z0-9]{1,10}-[A-Z0-9]{2,4}(?:-[A-Z0-9]{0,6})?)'  # Vendor Style ID (label opcjonalny)
    r'.+?'
    r'\s+(\d+)\s+EA\s+'                                      # Total Units
    r'(\d+\.\d{2})'                                          # Vendor First Cost
)

# ── NIKE/MACY'S parser ─────────────────────────────────────────────────────────
# Linia: 1 76F026-023-P3 NKB JUST DO IT SHORT SET BLACK 1 $9.00 300 $2,700.00 300 0
# Style: 6 znaków - 3 znaki - 2 znaki (label opcjonalny), np. 76F026-023-P3 lub 26M609-P6I-IP lub 76M336-U89
NIKE_LINE_PATTERN = re.compile(
    r'^\d{1,2}\s+'                                           # numer linii (1-99)
    r'([A-Z0-9]{6}-[A-Z0-9]{3}(?:-[A-Z0-9]{2})?)'          # Style# (6-3 lub 6-3-2)
    r'\s+.+?'                                                # opis
    r'\s+\$(\d+\.\d{2})'                                     # MBKS Cost
    r'\s+(\d+)'                                              # TOTAL UNITS
    r'\s+\$[\d,]+\.\d{2}'                                    # EXT COST (pomijamy)
)

def detect_format(text):
    """Rozpoznaje format pliku na podstawie treści."""
    if 'MBKS Cost' in text or 'MACYSNET' in text or 'macysnet' in text:
        return 'nike'
    if 'Vendor First Cost' in text or 'EA' in text and re.search(r'\d{10}[A-Z0-9]', text):
        return 'levis'
    return None

def extract_levis(pdf_file):
    rows = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split("\n"):
                line = line.strip()
                m = LEVIS_PATTERN.match(line)
                if m:
                    rows.append({
                        "Vendor Style ID": m.group(1),
                        "Total Units": int(m.group(2)),
                        "Vendor First Cost (USD)": float(m.group(3)),
                    })
    return rows

def extract_nike(pdf_file):
    rows = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split("\n"):
                line = line.strip()
                m = NIKE_LINE_PATTERN.match(line)
                if m:
                    rows.append({
                        "Vendor Style ID": m.group(1),
                        "Total Units": int(m.group(3)),
                        "Vendor First Cost (USD)": float(m.group(2)),
                    })
    return rows

def extract_po_data(pdf_file):
    # Wczytaj cały tekst żeby wykryć format
    full_text = ""
    pdf_file.seek(0)
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                full_text += t

    fmt = detect_format(full_text)
    pdf_file.seek(0)

    if fmt == 'nike':
        return extract_nike(pdf_file), 'Nike / Macy\'s Backstage'
    elif fmt == 'levis':
        return extract_levis(pdf_file), 'Levi\'s / Haddad'
    else:
        # Spróbuj obydwa parsery
        pdf_file.seek(0)
        rows = extract_levis(pdf_file)
        if rows:
            return rows, 'Levi\'s / Haddad'
        pdf_file.seek(0)
        rows = extract_nike(pdf_file)
        if rows:
            return rows, 'Nike / Macy\'s Backstage'
        return [], 'Nieznany'

def write_excel(df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="PO Data")
        ws = writer.sheets["PO Data"]

        # Wiersz z sumami na dole
        total_row = df.shape[0] + 2
        total_units = df["Total Units"].sum()
        total_cost = (df["Total Units"] * df["Vendor First Cost (USD)"]).sum()

        ws.cell(row=total_row, column=1).value = "TOTAL"
        ws.cell(row=total_row, column=2).value = total_units
        ws.cell(row=total_row, column=3).value = round(total_cost, 2)

        bold = Font(bold=True)
        fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        for col_idx in range(1, 4):
            cell = ws.cell(row=total_row, column=col_idx)
            cell.font = bold
            cell.fill = fill
            cell.alignment = Alignment(horizontal="center")

        for row in ws.iter_rows(min_row=2, max_row=total_row, min_col=3, max_col=3):
            for cell in row:
                cell.number_format = '#,##0.00'

        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = max_len + 4

    buffer.seek(0)
    return buffer


if uploaded_file is not None:
    with st.spinner("Przetwarzam PDF..."):
        try:
            rows, detected_format = extract_po_data(uploaded_file)

            if not rows:
                st.error("Nie udało się wyciągnąć danych. Sprawdź czy plik jest właściwym PO PDF.")
            else:
                df = pd.DataFrame(rows)

                st.info(f"📄 Wykryty format: **{detected_format}**")
                st.success(f"✅ Wyciągnięto **{len(df)}** pozycji z PDF.")
                st.dataframe(df, use_container_width=True)

                col1, col2 = st.columns(2)
                col1.metric("Łączna liczba jednostek", f"{df['Total Units'].sum():,}")
                col2.metric("Łączny koszt (USD)", f"${(df['Total Units'] * df['Vendor First Cost (USD)']).sum():,.2f}")

                buffer = write_excel(df)

                st.download_button(
                    label="⬇️ Pobierz Excel",
                    data=buffer,
                    file_name="po_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        except Exception as e:
            st.error(f"Błąd podczas przetwarzania: {e}")
            st.exception(e)
