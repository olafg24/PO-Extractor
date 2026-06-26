import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title="PO PDF → Excel", page_icon="📦", layout="centered")

st.title("📦 PO Extractor – Vendor Style ID / Units / Cost")
st.markdown("Wgraj plik PDF z Purchase Order (Levi's lub Nike/Macy's), a aplikacja wyciągnie dane i pozwoli pobrać Excel.")

uploaded_file = st.file_uploader("Wybierz plik PDF", type="pdf")

# ── LEVI'S parser ──────────────────────────────────────────────────────────────
LEVIS_PATTERN = re.compile(
    r'^\d{10}'
    r'([A-Z0-9]{1,10}-[A-Z0-9]{2,4}(?:-[A-Z0-9]{0,6})?)'
    r'.+?'
    r'\s+(\d+)\s+EA\s+'
    r'(\d+\.\d{2})'
)

# ── NIKE/MACY'S parser ─────────────────────────────────────────────────────────
NIKE_LINE_PATTERN = re.compile(
    r'^\d{1,2}\s+'
    r'([A-Z0-9]{6}-[A-Z0-9]{3}(?:-[A-Z0-9]{2})?)'
    r'\s+.+?'
    r'\s+\$(\d+\.\d{2})'
    r'\s+(\d+)'
    r'\s+\$[\d,]+\.\d{2}'
)

def detect_format(text):
    if 'MBKS Cost' in text or 'macysnet' in text.lower():
        return 'nike'
    if 'Vendor First Cost' in text or ('EA' in text and re.search(r'\d{10}[A-Z0-9]', text)):
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
        return extract_nike(pdf_file), "Nike / Macy's Backstage"
    elif fmt == 'levis':
        return extract_levis(pdf_file), "Levi's / Haddad"
    else:
        pdf_file.seek(0)
        rows = extract_levis(pdf_file)
        if rows:
            return rows, "Levi's / Haddad"
        pdf_file.seek(0)
        rows = extract_nike(pdf_file)
        if rows:
            return rows, "Nike / Macy's Backstage"
        return [], 'Nieznany'

def split_style_id(vendor_style_id):
    """
    Rozbija Vendor Style ID na części:
      76F026-023-P3  → style=76F026, color=023, label=P3
      76F026-023     → style=76F026, color=023, label=''
      76F026         → style=76F026, color='',  label=''
    """
    parts = vendor_style_id.split('-')
    style = parts[0] if len(parts) > 0 else ''
    color = parts[1] if len(parts) > 1 else ''
    label = parts[2] if len(parts) > 2 else ''
    return style, color, label

def write_excel_template(rows):
    """Wstawia dane do szablonu In-LineCart."""
    # Wczytaj szablon z dysku (Sheet1 jako docelowy arkusz)
    TEMPLATE_PATH = 'In-LineCart_template.xlsx'
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb['Sheet1']

    # Kolumny w szablonie: A=Div, B=Style#, C=Color, D=Label, E=Cost, F=SKU, G=Pcs
    start_row = 2
    for i, row in enumerate(rows):
        style, color, label = split_style_id(row['Vendor Style ID'])
        r = start_row + i
        ws.cell(row=r, column=1).value = ''           # Div - puste
        ws.cell(row=r, column=2).value = style        # Style#
        ws.cell(row=r, column=3).value = color        # Color
        ws.cell(row=r, column=4).value = label        # Label
        ws.cell(row=r, column=5).value = row['Vendor First Cost (USD)']  # Cost
        ws.cell(row=r, column=6).value = ''           # SKU - puste
        ws.cell(row=r, column=7).value = row['Total Units']              # Pcs

    # Wiersz TOTAL na dole
    total_row = start_row + len(rows)
    total_units = sum(r['Total Units'] for r in rows)
    total_cost = sum(r['Total Units'] * r['Vendor First Cost (USD)'] for r in rows)

    ws.cell(row=total_row, column=2).value = 'TOTAL'
    ws.cell(row=total_row, column=5).value = round(total_cost, 2)
    ws.cell(row=total_row, column=7).value = total_units

    bold = Font(bold=True)
    fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    center = Alignment(horizontal="center")
    for col_idx in [2, 5, 7]:
        cell = ws.cell(row=total_row, column=col_idx)
        cell.font = bold
        cell.fill = fill
        cell.alignment = center

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def write_excel_plain(rows):
    """Fallback - prosty Excel bez szablonu (gdy szablon nie jest wgrany)."""
    df = pd.DataFrame(rows)
    df['style'] = df['Vendor Style ID'].apply(lambda x: x.split('-')[0])
    df['color'] = df['Vendor Style ID'].apply(lambda x: x.split('-')[1] if len(x.split('-')) > 1 else '')
    df['label'] = df['Vendor Style ID'].apply(lambda x: x.split('-')[2] if len(x.split('-')) > 2 else '')

    out = pd.DataFrame({
        'Div': '',
        'Style#': df['style'],
        'Color': df['color'],
        'Label': df['label'],
        'Cost': df['Vendor First Cost (USD)'],
        'SKU': '',
        'Pcs': df['Total Units'],
    })

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="Sheet1")
        ws = writer.sheets["Sheet1"]

        total_row = len(out) + 2
        total_units = df['Total Units'].sum()
        total_cost = (df['Total Units'] * df['Vendor First Cost (USD)']).sum()

        ws.cell(row=total_row, column=2).value = 'TOTAL'
        ws.cell(row=total_row, column=5).value = round(total_cost, 2)
        ws.cell(row=total_row, column=7).value = total_units

        bold = Font(bold=True)
        fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        for col_idx in [2, 5, 7]:
            cell = ws.cell(row=total_row, column=col_idx)
            cell.font = bold
            cell.fill = fill
            cell.alignment = Alignment(horizontal="center")

        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = max_len + 4

    buffer.seek(0)
    return buffer


# ── UI ─────────────────────────────────────────────────────────────────────────

template_file = st.file_uploader(
    "📋 Wgraj szablon Excel (In-LineCart) — opcjonalnie",
    type=["xlsx"],
    help="Jeśli wgrasz szablon, dane zostaną wstawione do Sheet1. Bez szablonu aplikacja wygeneruje plik samodzielnie."
)

if uploaded_file is not None:
    with st.spinner("Przetwarzam PDF..."):
        try:
            rows, detected_format = extract_po_data(uploaded_file)

            if not rows:
                st.error("Nie udało się wyciągnąć danych. Sprawdź czy plik jest właściwym PO PDF.")
            else:
                # Podgląd rozbitych danych
                preview = []
                for row in rows:
                    style, color, label = split_style_id(row['Vendor Style ID'])
                    preview.append({
                        'Style#': style,
                        'Color': color,
                        'Label': label,
                        'Cost': row['Vendor First Cost (USD)'],
                        'Pcs': row['Total Units'],
                    })
                df_preview = pd.DataFrame(preview)

                st.info(f"📄 Wykryty format: **{detected_format}**")
                st.success(f"✅ Wyciągnięto **{len(rows)}** pozycji z PDF.")
                st.dataframe(df_preview, use_container_width=True)

                col1, col2 = st.columns(2)
                total_units = sum(r['Total Units'] for r in rows)
                total_cost = sum(r['Total Units'] * r['Vendor First Cost (USD)'] for r in rows)
                col1.metric("Łączna liczba jednostek", f"{total_units:,}")
                col2.metric("Łączny koszt (USD)", f"${total_cost:,.2f}")

                # Generuj Excel
                if template_file is not None:
                    # Zapisz szablon tymczasowo na dysk żeby openpyxl mógł go wczytać
                    template_bytes = template_file.read()
                    with open('In-LineCart_template.xlsx', 'wb') as f:
                        f.write(template_bytes)
                    buffer = write_excel_template(rows)
                    filename = "In-LineCart_filled.xlsx"
                else:
                    buffer = write_excel_plain(rows)
                    filename = "po_data.xlsx"

                st.download_button(
                    label="⬇️ Pobierz Excel",
                    data=buffer,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        except Exception as e:
            st.error(f"Błąd podczas przetwarzania: {e}")
            st.exception(e)
