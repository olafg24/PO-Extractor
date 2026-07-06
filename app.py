import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title="PO Extractor", page_icon="📦", layout="centered")

st.title("📦 PO Extractor")

# ── PARSERY ────────────────────────────────────────────────────────────────────

LEVIS_PATTERN = re.compile(
    r'^\d{10}'
    r'([A-Z0-9]{1,10}-[A-Z0-9]{2,4}(?:-[A-Z0-9]{0,6})?)'
    r'.+?'
    r'\s+(\d+)\s+EA\s+'
    r'(\d+\.\d{2})'
)

NIKE_LINE_PATTERN = re.compile(
    r'^\d{1,2}\s+'
    r'([A-Z0-9]{6}-[A-Z0-9]{3}(?:-{1,2}[A-Z0-9]{1,2})?)'  # -{1,2} obsługuje -- przed labelem
    r'\s+.+?'
    r'\s+\$(\d+\.\d{2})'
    r'\s+(\d+)'
    r'\s+\$[\d,]+\.\d{2}'
)

BABYLIST_PATTERN = re.compile(
    r'^(BL-\d+)'
    r'\s*\|\s*'
    r'([A-Z0-9]{5,8})\s*-\s*([A-Z0-9]{3})'
    r'(?:\s*-\s*([A-Z0-9]{2}))?'
    r'\s*\|.+?'
    r'\s+\$(\d+\.\d{2})'
    r'\s+(\d+)'
    r'\s+\$[\d,]+\.\d{2}$'
)

def extract_levis(pdf_file):
    rows = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text: continue
            for line in text.split("\n"):
                line = line.strip()
                m = LEVIS_PATTERN.match(line)
                if m:
                    rows.append({
                        "Vendor Style ID": m.group(1),
                        "Total Units": int(m.group(2)),
                        "Vendor First Cost (USD)": float(m.group(3)),
                        "SKU": "",
                    })
    return rows

def extract_nike(pdf_file):
    rows = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text: continue
            for line in text.split("\n"):
                line = line.strip()
                m = NIKE_LINE_PATTERN.match(line)
                if m:
                    rows.append({
                        "Vendor Style ID": m.group(1),
                        "Total Units": int(m.group(3)),
                        "Vendor First Cost (USD)": float(m.group(2)),
                        "SKU": "",
                    })
    return rows

def extract_babylist(pdf_file):
    rows = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text: continue
            for line in text.split("\n"):
                line = line.strip()
                m = BABYLIST_PATTERN.match(line)
                if m:
                    rows.append({
                        "Vendor Style ID": f"{m.group(2)}-{m.group(3)}" + (f"-{m.group(4)}" if m.group(4) else ""),
                        "Total Units": int(m.group(6)),
                        "Vendor First Cost (USD)": float(m.group(5)),
                        "SKU": m.group(1),
                    })
    return rows

ACCOUNT_PARSER = {
    "C810M": extract_levis,
    "B614M": extract_babylist,
    "M004M": extract_nike,
}

# ── HELPERS ────────────────────────────────────────────────────────────────────

def split_style_id(vendor_style_id):
    # Obsługuje zarówno pojedynczy myślnik (86F633-023-WH)
    # jak i podwójny myślnik (86M344-U90--V)
    import re as _re
    m = _re.match(r'^([A-Z0-9]+)-([A-Z0-9]+)-{1,2}([A-Z0-9]{1,2})$', vendor_style_id)
    if m:
        label = ('-' + m.group(3)) if len(m.group(3)) == 1 else m.group(3)
        return m.group(1), m.group(2), label
    parts = vendor_style_id.split('-')
    style = parts[0] if len(parts) > 0 else ''
    color = parts[1] if len(parts) > 1 else ''
    label = ''
    return style, color, label

def write_excel_template(rows):
    TEMPLATE_PATH = 'In-LineCart_template.xlsx'
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb['Sheet1']
    start_row = 2
    for i, row in enumerate(rows):
        style, color, label = split_style_id(row['Vendor Style ID'])
        r = start_row + i
        ws.cell(row=r, column=1).value = ''
        ws.cell(row=r, column=2).value = style
        ws.cell(row=r, column=3).value = color
        ws.cell(row=r, column=4).value = label
        ws.cell(row=r, column=5).value = row['Vendor First Cost (USD)']
        ws.cell(row=r, column=6).value = row.get('SKU', '')
        ws.cell(row=r, column=7).value = row['Total Units']

    total_row = start_row + len(rows)
    total_units = sum(r['Total Units'] for r in rows)
    total_cost = sum(r['Total Units'] * r['Vendor First Cost (USD)'] for r in rows)
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

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def write_excel_plain(rows):
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
        'SKU': df['SKU'] if 'SKU' in df.columns else '',
        'Pcs': df['Total Units'],
    })
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="Sheet1")
        ws = writer.sheets["Sheet1"]
        total_row = len(out) + 2
        ws.cell(row=total_row, column=2).value = 'TOTAL'
        ws.cell(row=total_row, column=5).value = round(df['Total Units'].mul(df['Vendor First Cost (USD)']).sum(), 2)
        ws.cell(row=total_row, column=7).value = int(df['Total Units'].sum())
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

# Inicjalizacja stanu
if "selected_account" not in st.session_state:
    st.session_state.selected_account = None

# Kafelki kont
st.markdown("### Wybierz konto")

ACCOUNTS = [
    {"name": "C810M", "color": "#1f4e79", "desc": "Century 21 / Haddad"},
    {"name": "B614M", "color": "#375623", "desc": "Babylist / Huggies"},
    {"name": "M004M", "color": "#7b2c2c", "desc": "Macy's Backstage"},
]

cols = st.columns(len(ACCOUNTS))
for col, acc in zip(cols, ACCOUNTS):
    with col:
        is_selected = st.session_state.selected_account == acc["name"]
        border = "4px solid #FFD700" if is_selected else "2px solid transparent"
        st.markdown(
            f"""
            <div style="
                background-color: {acc['color']};
                border: {border};
                border-radius: 12px;
                padding: 20px 10px;
                text-align: center;
                margin-bottom: 8px;
            ">
                <div style="color: white; font-size: 22px; font-weight: bold;">{acc['name']}</div>
                <div style="color: #ccc; font-size: 12px; margin-top: 4px;">{acc['desc']}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button(
            "✓ Wybrano" if is_selected else "Wybierz",
            key=f"btn_{acc['name']}",
            use_container_width=True,
        ):
            st.session_state.selected_account = acc["name"]
            st.rerun()

# Pokaż resztę UI tylko gdy konto jest wybrane
if st.session_state.selected_account:
    account = st.session_state.selected_account
    st.success(f"✅ Wybrane konto: **{account}**")
    st.divider()

    uploaded_file = st.file_uploader("Wybierz plik PDF", type="pdf")

    template_file = st.file_uploader(
        "📋 Wgraj szablon Excel (In-LineCart) — opcjonalnie",
        type=["xlsx"],
        help="Jeśli wgrasz szablon, dane zostaną wstawione do Sheet1."
    )

    if uploaded_file is not None:
        with st.spinner("Przetwarzam PDF..."):
            try:
                parser = ACCOUNT_PARSER[account]
                rows = parser(uploaded_file)

                if not rows:
                    st.error("Nie udało się wyciągnąć danych. Sprawdź czy wgrałeś właściwy plik dla tego konta.")
                else:
                    preview = []
                    for row in rows:
                        style, color, label = split_style_id(row['Vendor Style ID'])
                        preview.append({
                            'Style#': style,
                            'Color': color,
                            'Label': label,
                            'Cost': row['Vendor First Cost (USD)'],
                            'SKU': row.get('SKU', ''),
                            'Pcs': row['Total Units'],
                        })
                    df_preview = pd.DataFrame(preview)

                    st.success(f"✅ Wyciągnięto **{len(rows)}** pozycji z PDF.")
                    st.dataframe(df_preview, use_container_width=True)

                    col1, col2 = st.columns(2)
                    total_units = sum(r['Total Units'] for r in rows)
                    total_cost = sum(r['Total Units'] * r['Vendor First Cost (USD)'] for r in rows)
                    col1.metric("Łączna liczba jednostek", f"{total_units:,}")
                    col2.metric("Łączny koszt (USD)", f"${total_cost:,.2f}")

                    if template_file is not None:
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
else:
    st.info("👆 Najpierw wybierz konto powyżej.")
