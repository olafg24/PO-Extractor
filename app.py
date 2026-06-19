import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

st.set_page_config(page_title="PO PDF → Excel", page_icon="📦", layout="centered")

st.title("📦 PO Extractor – Vendor Style ID / Units / Cost")
st.markdown("Wgraj plik PDF z Purchase Order, a aplikacja wyciągnie dane i pozwoli pobrać Excel.")

uploaded_file = st.file_uploader("Wybierz plik PDF", type="pdf")

DATA_LINE_PATTERN = re.compile(
    r'^\d{10}'                                    # Internal Style ID (10 cyfr)
    r'([A-Z0-9]{1,10}-[A-Z0-9]{2,4}(?:-[A-Z0-9]{0,6})?)'  # Vendor Style ID (label opcjonalny)
    r'.+?'                                        # kolor + opis
    r'\s+(\d+)\s+EA\s+'                           # Total Units + EA
    r'(\d+\.\d{2})'                               # Vendor First Cost
)

def extract_po_data(pdf_file):
    rows = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split("\n"):
                line = line.strip()
                m = DATA_LINE_PATTERN.match(line)
                if m:
                    rows.append({
                        "Vendor Style ID": m.group(1),
                        "Total Units": int(m.group(2)),
                        "Vendor First Cost (USD)": float(m.group(3)),
                    })
    return rows


if uploaded_file is not None:
    with st.spinner("Przetwarzam PDF..."):
        try:
            rows = extract_po_data(uploaded_file)

            if not rows:
                st.error("Nie udało się wyciągnąć danych. Sprawdź czy plik jest właściwym PO PDF.")
            else:
                df = pd.DataFrame(rows)

                st.success(f"✅ Wyciągnięto **{len(df)}** pozycji z PDF.")
                st.dataframe(df, use_container_width=True)

                col1, col2 = st.columns(2)
                col1.metric("Łączna liczba jednostek", f"{df['Total Units'].sum():,}")
                col2.metric("Łączny koszt (USD)", f"${(df['Total Units'] * df['Vendor First Cost (USD)']).sum():,.2f}")

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name="PO Data")
                    ws = writer.sheets["PO Data"]

                    # Wiersz z sumami na dole
                    total_row = df.shape[0] + 2  # +1 nagłówek, +1 bo Excel 1-indexed
                    total_units = df["Total Units"].sum()
                    total_cost = (df["Total Units"] * df["Vendor First Cost (USD)"]).sum()

                    ws.cell(row=total_row, column=1).value = "TOTAL"
                    ws.cell(row=total_row, column=2).value = total_units
                    ws.cell(row=total_row, column=3).value = round(total_cost, 2)

                    # Pogrubienie wiersza z sumami
                    from openpyxl.styles import Font, PatternFill, Alignment
                    bold = Font(bold=True)
                    fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                    for col_idx in range(1, 4):
                        cell = ws.cell(row=total_row, column=col_idx)
                        cell.font = bold
                        cell.fill = fill
                        cell.alignment = Alignment(horizontal="center")

                    # Format kolumny z kosztem jako waluta
                    for row in ws.iter_rows(min_row=2, max_row=total_row, min_col=3, max_col=3):
                        for cell in row:
                            cell.number_format = '#,##0.00'

                    # Auto-szerokość kolumn
                    for col in ws.columns:
                        max_len = max(len(str(cell.value or "")) for cell in col)
                        ws.column_dimensions[col[0].column_letter].width = max_len + 4

                buffer.seek(0)
                st.download_button(
                    label="⬇️ Pobierz Excel",
                    data=buffer,
                    file_name="po_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        except Exception as e:
            st.error(f"Błąd podczas przetwarzania: {e}")
            st.exception(e)
