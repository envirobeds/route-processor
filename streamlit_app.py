import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import datetime, re, io

st.set_page_config(page_title='Route Processor', page_icon='🚛', layout='centered')

st.markdown("""
<style>
    .main { max-width: 700px; }
    .stButton>button { background-color: #1D3557; color: white; width: 100%; padding: 0.6rem; font-size: 1.1rem; border-radius: 8px; border: none; }
    .stButton>button:hover { background-color: #457B9D; }
    .success-box { background: #EAF3DE; border-radius: 8px; padding: 1rem 1.5rem; margin-top: 1rem; }
</style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='color:#1D3557'>🚛 Route Processor</h1>", unsafe_allow_html=True)
st.markdown("<p style='color:#666'>Upload your route CSV and download polished Excel files instantly.</p>", unsafe_allow_html=True)
st.divider()

# ── Colours ───────────────────────────────────────────────────────────────────
NAVY='1D3557'; MID_BLUE='457B9D'; LIGHT_BLUE='A8D8EA'; ALT_ROW='EBF5FB'
WHITE='FFFFFF'; DARK_TEXT='1A1A2E'; TOTAL_BG='D4E6F1'
ORANGE='C75B00'; LIGHT_ORANGE='FEF0E7'; ORANGE_MID='E8874A'; ORANGE_HDR='FDDCCA'

def fill(hex): return PatternFill('solid', fgColor=hex)
def font(bold=False, sz=10, color=DARK_TEXT, italic=False):
    return Font(name='Calibri', bold=bold, size=sz, color=color, italic=italic)
def thin_border(): return Border(bottom=Side(style='thin', color='D5E8F5'))

COL_CAPS = {
    'Ref ID':10,'Date':12,'Date Booked':12,'Date Collected':14,
    'Suburb':14,'Qty Booked':9,'Qty Collected':11,
    'Collection Notes':45,'Notes':45,'Photo URL':45,'Tracking Link':40,
}
def col_width(series, header):
    longest = max(series.astype(str).str.len().max(), len(header)) + 2
    return int(COL_CAPS.get(header, longest))

def clean_date(val):
    val = str(val).strip().split(' ')[0]
    try: return pd.to_datetime(val, dayfirst=True)
    except: return pd.NaT

def extract_suburb_from_address(address):
    address = str(address)
    m = re.search(r',\s*([A-Za-z\s]+),\s*Randwick City Council', address, re.IGNORECASE)
    if m: return m.group(1).strip().title()
    m = re.search(r'\b([A-Z]+(?:\s+[A-Z]+)?)\s*,\s*NSW', address)
    if m: return m.group(1).strip().title()
    m = re.search(r',\s*([A-Z][A-Z\s]+)\s*$', address, re.IGNORECASE)
    if m: return m.group(1).strip().title()
    return ''

def extract_suburb(row):
    if pd.isna(row['suburb']) or str(row['suburb']).strip() == '':
        return extract_suburb_from_address(row['address'])
    return str(row['suburb']).strip().title()

col_map_oncall = {
    'seller_order_id':'Ref ID','date_booked':'Date','suburb':'Suburb',
    'address':'Address','notes':'Collection Notes','qty_booked':'Qty Booked',
    'driver_provided_recipient_notes':'Qty Collected',
    'photo_url':'Photo URL','tracking_url':'Tracking Link',
}
col_map_booked = {
    'seller_order_id':'Ref ID','date_booked':'Date Booked',
    'address':'Address','notes':'Notes',
    'driver_provided_recipient_notes':'Qty Collected',
    'photo_url':'Photo URL','tracking_url':'Tracking Link',
}

def build_oncall(ws, df_data, route_name):
    df_out = df_data[list(col_map_oncall.keys())].rename(columns=col_map_oncall)
    headers = list(col_map_oncall.values())
    num_cols = len(headers)

    ws.row_dimensions[1].height = 28
    ws.merge_cells(f'A1:{get_column_letter(num_cols)}1')
    c = ws['A1']
    c.value = f'  {route_name}  —  On-Call Mattress Collection'
    c.fill = fill(NAVY); c.font = Font(name='Calibri', bold=True, size=14, color=WHITE)
    c.alignment = Alignment(horizontal='left', vertical='center')

    ws.row_dimensions[2].height = 16
    ws.merge_cells(f'A2:{get_column_letter(num_cols)}2')
    c = ws['A2']
    c.value = (f'  Generated: {datetime.date.today().strftime("%d/%m/%Y")}   |   '
               f'Total stops: {len(df_out)}   |   Qty collected: {int(df_out["Qty Collected"].sum())}')
    c.fill = fill(MID_BLUE); c.font = Font(name='Calibri', size=9, color=WHITE, italic=True)
    c.alignment = Alignment(horizontal='left', vertical='center')

    ws.row_dimensions[3].height = 18
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=3, column=ci, value=h)
        c.fill = fill(LIGHT_BLUE); c.font = font(bold=True, sz=10, color=NAVY)
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = Border(bottom=Side(style='medium', color=MID_BLUE))

    for ri, row in df_out.iterrows():
        excel_row = ri + 4
        ws.row_dimensions[excel_row].height = 13
        row_fill = fill(ALT_ROW) if ri % 2 == 1 else fill(WHITE)
        for ci, (col, val) in enumerate(row.items(), 1):
            c = ws.cell(row=excel_row, column=ci, value=val)
            c.fill = row_fill; c.font = font(sz=9); c.border = thin_border()
            if col in ('Qty Booked','Qty Collected','Ref ID','Date'):
                c.alignment = Alignment(horizontal='center', vertical='center')
            else:
                c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)

    total_row = len(df_out) + 4
    ws.row_dimensions[total_row].height = 16
    collected_col = headers.index('Qty Collected') + 1
    for ci in range(1, num_cols + 1):
        c = ws.cell(row=total_row, column=ci)
        c.fill = fill(TOTAL_BG); c.border = Border(top=Side(style='medium', color=MID_BLUE))
        if ci == 1:
            c.value = 'TOTALS'; c.font = font(bold=True, sz=10, color=NAVY)
            c.alignment = Alignment(horizontal='left', vertical='center')
        elif ci == collected_col:
            c.value = f'=SUM({get_column_letter(ci)}4:{get_column_letter(ci)}{total_row-1})'
            c.font = font(bold=True, sz=10, color=NAVY)
            c.alignment = Alignment(horizontal='center', vertical='center')

    for ci, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(ci)].width = col_width(df_out[h], h)
    ws.freeze_panes = 'A4'
    ws.auto_filter.ref = f'A3:{get_column_letter(num_cols)}3'


def build_booked(ws, df_data, route_name, date_collected):
    df_out = df_data[list(col_map_booked.keys())].rename(columns=col_map_booked).copy()
    df_out.insert(2, 'Date Collected', date_collected)
    headers = list(df_out.columns)
    num_cols = len(headers)

    ws.row_dimensions[1].height = 28
    ws.merge_cells(f'A1:{get_column_letter(num_cols)}1')
    c = ws['A1']
    c.value = f'  {route_name}  —  Booked Mattress Collection (WM28)'
    c.fill = fill(ORANGE); c.font = Font(name='Calibri', bold=True, size=14, color=WHITE)
    c.alignment = Alignment(horizontal='left', vertical='center')

    ws.row_dimensions[2].height = 16
    ws.merge_cells(f'A2:{get_column_letter(num_cols)}2')
    c = ws['A2']
    c.value = (f'  Generated: {datetime.date.today().strftime("%d/%m/%Y")}   |   '
               f'Total stops: {len(df_out)}   |   Qty collected: {int(df_out["Qty Collected"].sum())}')
    c.fill = fill(ORANGE_MID); c.font = Font(name='Calibri', size=9, color=WHITE, italic=True)
    c.alignment = Alignment(horizontal='left', vertical='center')

    ws.row_dimensions[3].height = 18
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=3, column=ci, value=h)
        c.fill = fill(ORANGE_HDR); c.font = font(bold=True, sz=10, color=ORANGE)
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = Border(bottom=Side(style='medium', color=ORANGE_MID))

    for ri, row in df_out.iterrows():
        excel_row = ri + 4
        ws.row_dimensions[excel_row].height = 13
        row_fill = fill(LIGHT_ORANGE) if ri % 2 == 1 else fill(WHITE)
        for ci, (col, val) in enumerate(row.items(), 1):
            c = ws.cell(row=excel_row, column=ci, value=val)
            c.fill = row_fill; c.font = font(sz=9); c.border = thin_border()
            if col in ('Ref ID','Date Booked','Date Collected','Qty Collected'):
                c.alignment = Alignment(horizontal='center', vertical='center')
            else:
                c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)

    total_row = len(df_out) + 4
    ws.row_dimensions[total_row].height = 16
    collected_col = headers.index('Qty Collected') + 1
    for ci in range(1, num_cols + 1):
        c = ws.cell(row=total_row, column=ci)
        c.fill = fill(TOTAL_BG); c.border = Border(top=Side(style='medium', color=ORANGE_MID))
        if ci == 1:
            c.value = 'TOTALS'; c.font = font(bold=True, sz=10, color=ORANGE)
            c.alignment = Alignment(horizontal='left', vertical='center')
        elif ci == collected_col:
            c.value = f'=SUM({get_column_letter(ci)}4:{get_column_letter(ci)}{total_row-1})'
            c.font = font(bold=True, sz=10, color=ORANGE)
            c.alignment = Alignment(horizontal='center', vertical='center')

    for ci, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(ci)].width = col_width(df_out[h], h)
    ws.freeze_panes = 'A4'
    ws.auto_filter.ref = f'A3:{get_column_letter(num_cols)}3'


def process_csv(df):
    df = df.drop_duplicates(subset=['seller_order_id', 'address'])
    df['date_booked_dt'] = df['date_booked'].apply(clean_date)
    route_date = df['date_booked_dt'].dropna().mode()[0]
    df['date_booked_dt'] = df['date_booked_dt'].fillna(route_date)
    df['date_booked'] = df['date_booked_dt'].apply(lambda x: x.strftime('%d/%m/%Y'))
    df['suburb'] = df.apply(extract_suburb, axis=1)

    route_name = df['route'].iloc[0]
    route_clean = re.sub(r'^[A-Za-z]+,\s*', '', route_name)
    oncall_filename = f'{route_name} - On Call'
    booked_filename = f'Booked {route_clean}'
    route_date_str = route_date.strftime('%d/%m/%Y')

    df_wm28  = df[df['products'] == 'WM28'].copy()
    df_bulky = df[df['products'] == 'Bulky Mattress'].copy()
    df_blank = df[df['products'].isna() | (df['products'].str.strip() == '')].copy()

    df_wm28  = df_wm28.sort_values(['date_booked_dt','address']).reset_index(drop=True)
    df_bulky = df_bulky.sort_values(['date_booked_dt','suburb','address']).reset_index(drop=True)
    df_blank = df_blank.sort_values(['date_booked_dt','suburb']).reset_index(drop=True)
    df_oncall = pd.concat([df_bulky, df_blank], ignore_index=True)

    # On Call — save to bytes buffer
    wb1 = Workbook()
    ws1 = wb1.active; ws1.title = 'On Call'
    build_oncall(ws1, df_oncall, route_name)
    buf1 = io.BytesIO()
    wb1.save(buf1)
    buf1.seek(0)

    # Booked — save to bytes buffer
    wb2 = Workbook()
    ws2 = wb2.active; ws2.title = 'Booked'
    build_booked(ws2, df_wm28, route_name, route_date_str)
    buf2 = io.BytesIO()
    wb2.save(buf2)
    buf2.seek(0)

    return buf1, buf2, oncall_filename, booked_filename, len(df_oncall), len(df_wm28)


# ── UI ────────────────────────────────────────────────────────────────────────
uploaded = st.file_uploader('Upload your route CSV', type=['csv'])

if uploaded:
    df = pd.read_csv(uploaded)
    route_name = df['route'].iloc[0] if 'route' in df.columns else 'Unknown Route'

    col1, col2, col3 = st.columns(3)
    col1.metric('Route', route_name.split(',')[-1].strip())
    col2.metric('Total Stops', len(df))
    col3.metric('Products', df['products'].value_counts().to_dict().__str__().replace('{','').replace('}','').replace("'",''))

    st.divider()

    if st.button('Generate Excel Files'):
        with st.spinner('Processing...'):
            try:
                buf1, buf2, f1, f2, n1, n2 = process_csv(df)

                st.success('Done! Download your files below.')

                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("**📋 On Call**")
                    st.download_button(
                        label=f'Download ({n1} rows)',
                        data=buf1,
                        file_name=f'{f1}.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        use_container_width=True
                    )
                with col2:
                    st.markdown("**📋 Booked (WM28)**")
                    st.download_button(
                        label=f'Download ({n2} rows)',
                        data=buf2,
                        file_name=f'{f2}.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        use_container_width=True
                    )
            except Exception as e:
                st.error(f'Error: {e}')
else:
    st.info('👆 Upload a CSV file to get started')
