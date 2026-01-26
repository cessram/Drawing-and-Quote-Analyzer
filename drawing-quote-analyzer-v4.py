import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import re
from datetime import datetime

try:
    import pdfplumber
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False

st.set_page_config(page_title="Drawing Quote Analyzer", page_icon="📊", layout="wide")

st.markdown("""
<style>
    .main-header { font-size: 2.5rem; font-weight: bold; color: #1f4e79; }
    .sub-header { font-size: 1.2rem; color: #666; }
    div[data-testid="stMetricValue"] { font-size: 1.6rem; }
</style>
""", unsafe_allow_html=True)

DEFAULT_SUPPLIER_CODES = {
    1: "Owner Supply / Owner Install",
    2: "Owner Supply / Owner Install (Special)",
    3: "Owner Supply / Owner Install (Other)",
    4: "Owner Supply / Vendor Install",
    5: "Contractor Supply / Contractor Install",
    6: "Contractor Supply / Vendor Install",
    7: "Owner Supply / Contractor Install",
    8: "Existing / Relocated"
}

# Initialize session state
for key in ['drawing_data', 'drawing_df', 'drawing_filename']:
    if key not in st.session_state:
        st.session_state[key] = None
for key in ['quotes_data', 'quote_dfs', 'quote_mappings', 'column_mapping']:
    if key not in st.session_state:
        st.session_state[key] = {}
if 'supplier_codes' not in st.session_state:
    st.session_state.supplier_codes = DEFAULT_SUPPLIER_CODES.copy()
if 'use_categories' not in st.session_state:
    st.session_state.use_categories = True

def clean_dataframe_columns(df):
    df = df.copy()
    new_cols = []
    seen = {}
    for i, c in enumerate(df.columns):
        col_name = str(c).strip() if pd.notna(c) and str(c).strip() != '' else f'Column_{i}'
        if col_name in seen:
            seen[col_name] += 1
            col_name = f"{col_name}_{seen[col_name]}"
        else:
            seen[col_name] = 0
        new_cols.append(col_name)
    df.columns = new_cols
    df = df.dropna(how='all')
    return df

def extract_quote_from_pdf_text(uploaded_file):
    if not PDF_SUPPORT:
        return None
    uploaded_file.seek(0)
    items = []
    try:
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                lines = text.split('\n')
                for line in lines:
                    line = line.strip()
                    if not line:
                        continue
                    skip_patterns = ['Canadian Restaurant Supply', 'Bird Construc', 'Page ', 'FWG LTC', 'Quote valid', 'Inspections:', 'Item Qty Description', 'ITEM TOTAL:', 'Merchandise', 'GST', 'Tax', 'Total']
                    if any(p in line for p in skip_patterns):
                        continue
                    
                    range_nic = re.match(r'^(\d+)[-–](\d+)\s+NIC\s*$', line, re.IGNORECASE)
                    if range_nic:
                        start, end = int(range_nic.group(1)), int(range_nic.group(2))
                        for num in range(start, end + 1):
                            items.append({'Item': str(num), 'Qty': '', 'Description': 'NIC', 'Sell': '', 'Sell_Total': ''})
                        continue
                    
                    single_nic = re.match(r'^(\d+)\s+NIC\s*$', line, re.IGNORECASE)
                    if single_nic:
                        items.append({'Item': single_nic.group(1), 'Qty': '', 'Description': 'NIC', 'Sell': '', 'Sell_Total': ''})
                        continue
                    
                    item_full = re.match(r'^(\d+)\s+(\d+)\s*ea\s+([A-Z][A-Z0-9\s,./\-&\(\)\'\"]+?)\s+\$?([\d,]+\.?\d*)\s+\$?([\d,]+\.?\d*)\s*$', line, re.IGNORECASE)
                    if item_full:
                        items.append({'Item': item_full.group(1), 'Qty': f"{item_full.group(2)} ea", 'Description': item_full.group(3).strip(), 'Sell': item_full.group(4), 'Sell_Total': item_full.group(5)})
                        continue
        if items:
            return pd.DataFrame(items)
    except Exception as e:
        st.warning(f"Text extraction error: {e}")
    return None

def parse_pdf_tables_for_quote(uploaded_file):
    if not PDF_SUPPORT:
        return None
    uploaded_file.seek(0)
    all_rows = []
    try:
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table or len(table) < 1:
                        continue
                    header_idx = -1
                    for i, row in enumerate(table):
                        row_text = ' '.join([str(c).lower() if c else '' for c in row])
                        if 'item' in row_text and ('qty' in row_text or 'description' in row_text or 'sell' in row_text):
                            header_idx = i
                            break
                    if header_idx >= 0:
                        headers = [str(h).strip() if h else f'Col_{j}' for j, h in enumerate(table[header_idx])]
                        for row in table[header_idx + 1:]:
                            if row and any(cell for cell in row if cell and str(cell).strip()):
                                row_dict = {headers[j]: cell for j, cell in enumerate(row) if j < len(headers)}
                                all_rows.append(row_dict)
                    else:
                        for row in table:
                            if row and len(row) >= 2:
                                first_cell = str(row[0]).strip() if row[0] else ''
                                if re.match(r'^\d+(-\d+)?$', first_cell):
                                    all_rows.append({'Item': row[0], 'Qty': row[1] if len(row) > 1 else '', 'Description': row[2] if len(row) > 2 else '', 'Sell': row[3] if len(row) > 3 else '', 'Sell Total': row[4] if len(row) > 4 else ''})
        if all_rows:
            df = pd.DataFrame(all_rows)
            return [clean_dataframe_columns(df)]
    except Exception as e:
        st.warning(f"PDF table extraction error: {e}")
    return None

def parse_pdf_tables(uploaded_file):
    if not PDF_SUPPORT:
        return None
    uploaded_file.seek(0)
    all_tables = []
    try:
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and len(table) > 1:
                        headers = [str(h).strip() if h else f'Col_{i}' for i, h in enumerate(table[0])]
                        df = pd.DataFrame(table[1:], columns=headers)
                        df = clean_dataframe_columns(df)
                        if len(df) > 0:
                            all_tables.append(df)
    except Exception as e:
        st.warning(f"Could not extract tables from PDF: {e}")
        return None
    return all_tables if all_tables else None

def parse_excel_file(uploaded_file):
    try:
        uploaded_file.seek(0)
        xl = pd.ExcelFile(uploaded_file)
        dfs = [clean_dataframe_columns(pd.read_excel(xl, sheet_name=name)) for name in xl.sheet_names]
        return [df for df in dfs if len(df) > 0]
    except Exception as e:
        st.warning(f"Could not read Excel file: {e}")
        return None

def parse_csv_file(uploaded_file):
    try:
        uploaded_file.seek(0)
        df = clean_dataframe_columns(pd.read_csv(uploaded_file))
        return [df] if len(df) > 0 else None
    except Exception as e:
        st.warning(f"Could not read CSV file: {e}")
        return None

def parse_uploaded_file(uploaded_file, file_type='drawing'):
    ext = uploaded_file.name.split('.')[-1].lower()
    if ext == 'pdf':
        if file_type == 'quote':
            text_df = extract_quote_from_pdf_text(uploaded_file)
            if text_df is not None and len(text_df) > 0:
                return [text_df]
            return parse_pdf_tables_for_quote(uploaded_file)
        return parse_pdf_tables(uploaded_file)
    elif ext in ['xlsx', 'xls']:
        return parse_excel_file(uploaded_file)
    elif ext == 'csv':
        return parse_csv_file(uploaded_file)
    return None

def auto_detect_columns(df, file_type='drawing'):
    cols_lower = {c: c.lower().strip() for c in df.columns}
    if file_type == 'drawing':
        patterns = {'no': ['no', 'no.', 'item', 'item #', 'number', '#', 'id'], 'description': ['description', 'desc', 'equipment', 'name', 'material'], 'qty': ['qty', 'qty.', 'quantity', 'count'], 'category': ['category', 'cat', 'supplier code', 'code', 'type'], 'equip_num': ['equipment number', 'equip num', 'model', 'part no']}
    else:
        patterns = {'no': ['item', 'no', 'no.', 'item #', 'number', '#', 'id', 'line'], 'description': ['description', 'desc', 'equipment', 'name', 'material', 'product'], 'qty': ['qty', 'qty.', 'quantity', 'count', 'ea'], 'unit_price': ['sell', 'unit price', 'price', 'rate', 'unit cost', 'each', 'unit'], 'total_price': ['sell_total', 'sell total', 'total', 'total price', 'ext price', 'extended', 'amount']}
    found = {}
    for key, opts in patterns.items():
        for col, col_low in cols_lower.items():
            if col_low in opts or any(opt in col_low for opt in opts):
                found[key] = col
                break
    return found

def clean_numeric(val):
    if pd.isna(val):
        return None
    val_str = re.sub(r'[,$]', '', str(val).strip())
    val_str = re.sub(r'[^\d.\-]', '', val_str)
    try:
        return float(val_str) if val_str else None
    except:
        return None

def parse_qty_value(val):
    if pd.isna(val):
        return 1
    val_str = str(val).strip().lower()
    if not val_str or val_str in ('nan', 'none', ''):
        return 1
    match = re.search(r'(\d+)\s*ea', val_str)
    if match:
        return int(match.group(1))
    num = clean_numeric(val)
    return int(num) if num and num > 0 else 1

def extract_drawing_data(df, col_map):
    items = []
    no_col, desc_col = col_map.get('no'), col_map.get('description')
    qty_col, cat_col, equip_col = col_map.get('qty'), col_map.get('category'), col_map.get('equip_num')
    if not no_col or not desc_col:
        return None
    for idx, row in df.iterrows():
        try:
            no_val = str(row.get(no_col, '')).strip()
            desc_val = str(row.get(desc_col, '')).strip()
            if not no_val or no_val.lower() in ('nan', '', 'no', 'no.', 'item', 'none'):
                continue
            if not desc_val or desc_val.lower() in ('nan', '', 'description', 'none'):
                continue
            qty = int(clean_numeric(row.get(qty_col, 1)) or 1) if qty_col else 1
            cat = int(clean_numeric(row.get(cat_col, ''))) if cat_col and clean_numeric(row.get(cat_col, '')) else None
            equip_num = str(row.get(equip_col, '')).strip() if equip_col else '-'
            equip_num = equip_num if equip_num and equip_num.lower() not in ('nan', '', '-', 'none') else '-'
            items.append({'No': no_val, 'Equip_Num': equip_num, 'Description': desc_val, 'Qty': qty, 'Category': cat})
        except:
            continue
    return items if items else None

def extract_quote_data(df, col_map, source_file):
    items = []
    no_col, desc_col, qty_col = col_map.get('no'), col_map.get('description'), col_map.get('qty')
    unit_col, total_col = col_map.get('unit_price'), col_map.get('total_price')
    for idx, row in df.iterrows():
        try:
            no_val = str(row.get(no_col, '')).strip() if no_col and no_col in df.columns else ''
            if not no_val or no_val.lower() in ('nan', 'none', 'item', ''):
                continue
            desc_val = str(row.get(desc_col, '')).strip() if desc_col and desc_col in df.columns else ''
            if desc_val.lower() in ('nan', 'none', 'description'):
                desc_val = ''
            qty_raw = str(row.get(qty_col, '')).strip() if qty_col and qty_col in df.columns else ''
            is_nic = desc_val.upper() == 'NIC' or 'NIC' in desc_val.upper() or qty_raw.upper() == 'NIC'
            if is_nic:
                desc_val = 'NIC'
            qty = parse_qty_value(row.get(qty_col, '')) if qty_col and qty_col in df.columns and not is_nic else 1
            unit_price = clean_numeric(row.get(unit_col, 0)) or 0 if unit_col and unit_col in df.columns else 0
            total_price = clean_numeric(row.get(total_col, 0)) or 0 if total_col and total_col in df.columns else 0
            if total_price == 0 and unit_price > 0:
                total_price = unit_price * qty
            if unit_price == 0 and total_price > 0 and qty > 0:
                unit_price = total_price / qty
            
            range_match = re.match(r'^(\d+)[-–](\d+)$', no_val)
            if range_match:
                start, end = int(range_match.group(1)), int(range_match.group(2))
                for num in range(start, end + 1):
                    items.append({'Item_No': str(num), 'Description': 'NIC', 'Qty': 1, 'Unit_Price': 0, 'Total_Price': 0, 'Is_NIC': True, 'Source_File': source_file})
            else:
                items.append({'Item_No': no_val, 'Description': desc_val if desc_val else '-', 'Qty': qty, 'Unit_Price': unit_price, 'Total_Price': total_price, 'Is_NIC': is_nic, 'Source_File': source_file})
        except:
            continue
    return items

def match_items(drawing_no, quotes):
    drawing_no_clean = str(drawing_no).strip().lower()
    for q in quotes:
        if str(q.get('Item_No', '')).strip().lower() == drawing_no_clean:
            return q
    try:
        drawing_num = int(re.sub(r'[^0-9]', '', drawing_no_clean))
        for q in quotes:
            try:
                quote_num = int(re.sub(r'[^0-9]', '', str(q.get('Item_No', '')).strip()))
                if drawing_num == quote_num:
                    return q
            except:
                pass
    except:
        pass
    return None

def analyze_data(drawing_items, quotes, use_categories=True, supplier_codes=None):
    if supplier_codes is None:
        supplier_codes = DEFAULT_SUPPLIER_CODES
    analysis = []
    for item in drawing_items:
        match = match_items(item['No'], quotes)
        cat = item.get('Category')
        desc_upper = item.get('Description', '').upper()
        
        if desc_upper in ('SPARE', '-', 'N/A') or 'SPARE' in desc_upper:
            status, issue = "N/A", "Spare Item"
        elif use_categories and cat in [1, 2, 3]:
            status, issue = "Owner Supply", f"{supplier_codes.get(cat, 'Owner handles')} - Excluded"
        elif use_categories and cat == 8:
            status, issue = "Existing", "Existing/Relocated Equipment"
        elif match:
            if match.get('Is_NIC'):
                status, issue = "🚫 NIC", "Not In Contract"
            elif match['Total_Price'] == 0 and match['Qty'] == item['Qty']:
                status, issue = "⚡ Included", "Included in system pricing"
            elif match['Qty'] == item['Qty']:
                status, issue = "✓ Quoted", None
            else:
                status, issue = "⚠ Qty Mismatch", f"Drawing: {item['Qty']}, Quote: {match['Qty']}"
        else:
            if use_categories and cat == 7:
                status, issue = "⚠ Needs Install", "Owner supplies - needs installation quote"
            elif use_categories and cat in [5, 6]:
                status, issue = "❌ MISSING", f"CRITICAL - Contractor Supply (Code {cat}) not quoted!"
            elif use_categories and cat == 4:
                status, issue = "❌ Missing", "Owner Supply / Vendor Install - Not quoted"
            else:
                status, issue = "❌ Missing", "Not found in quotes"
        
        analysis.append({
            'Drawing_No': item['No'], 'Equip_Num': item.get('Equip_Num', '-'), 'Description': item['Description'],
            'Drawing_Qty': item['Qty'], 'Category': cat,
            'Category_Desc': supplier_codes.get(cat, '-') if cat and use_categories else '-',
            'Quote_Item_No': match['Item_No'] if match else '-',
            'Quote_Qty': match['Qty'] if match else 0,
            'Unit_Price': match['Unit_Price'] if match else 0,
            'Total_Price': match['Total_Price'] if match and not match.get('Is_NIC') else 0,
            'Quote_Source': match['Source_File'] if match else '-',
            'Status': status, 'Issue': issue
        })
    return pd.DataFrame(analysis)

def get_supplier_code_summary(drawing_items, results_df, supplier_codes):
    schedule_df = pd.DataFrame(drawing_items)
    summary_data = []
    for code in range(1, 9):
        schedule_items = schedule_df[schedule_df['Category'] == code]
        line_items = len(schedule_items)
        total_qty = schedule_items['Qty'].sum() if not schedule_items.empty else 0
        code_results = results_df[results_df['Category'] == code]
        
        quoted_items = len(code_results[code_results['Status'].isin(['✓ Quoted', '⚡ Included'])])
        missing_items = len(code_results[code_results['Status'].str.contains('MISSING|Missing', case=False, na=False)])
        nic_items = len(code_results[code_results['Status'].str.contains('NIC', na=False)])
        mismatch_items = len(code_results[code_results['Status'] == '⚠ Qty Mismatch'])
        needs_install = len(code_results[code_results['Status'] == '⚠ Needs Install'])
        quoted_value = code_results[code_results['Status'].isin(['✓ Quoted', '⚡ Included', '⚠ Qty Mismatch'])]['Total_Price'].sum()
        
        if code in [1, 2, 3]:
            coverage = "N/A (Owner Supply)"
        elif code == 8:
            coverage = "N/A (Existing)"
        elif line_items > 0:
            coverage = f"{(quoted_items / line_items) * 100:.1f}%"
        else:
            coverage = "N/A"
        
        summary_data.append({
            "Code": code, "Description": supplier_codes[code], "Schedule Items": line_items,
            "Total Qty": int(total_qty), "Quoted": quoted_items, "Missing": missing_items,
            "NIC": nic_items, "Mismatch": mismatch_items, "Needs Install": needs_install,
            "Quoted Value": quoted_value, "Quote Required": "No" if code in [1, 2, 3, 8] else "Yes",
            "Coverage": coverage
        })
    return pd.DataFrame(summary_data)

def create_excel_report(drawing_items, results_df, quotes, supplier_summary_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        quoted = len(results_df[results_df['Status'] == '✓ Quoted'])
        included = len(results_df[results_df['Status'] == '⚡ Included'])
        missing = len(results_df[results_df['Status'].str.contains('MISSING|Missing', case=False, na=False)])
        nic = len(results_df[results_df['Status'].str.contains('NIC', na=False)])
        mismatch = len(results_df[results_df['Status'] == '⚠ Qty Mismatch'])
        needs_install = len(results_df[results_df['Status'] == '⚠ Needs Install'])
        
        summary = pd.DataFrame({
            "Metric": ["Report Date", "Total Items", "✓ Quoted", "⚡ Included", "❌ MISSING", "🚫 NIC", "⚠ Needs Install", "⚠ Mismatch", "Total Quoted Value"],
            "Value": [datetime.now().strftime("%Y-%m-%d %H:%M"), len(results_df), quoted, included, missing, nic, needs_install, mismatch, f"${results_df['Total_Price'].sum():,.2f}"]
        })
        summary.to_excel(writer, sheet_name='Executive Summary', index=False)
        
        sup_disp = supplier_summary_df.copy()
        sup_disp['Quoted Value'] = sup_disp['Quoted Value'].apply(lambda x: f"${x:,.2f}")
        sup_disp.to_excel(writer, sheet_name='Supplier Code Summary', index=False)
        
        results_df.to_excel(writer, sheet_name='Full Analysis', index=False)
        results_df[results_df['Status'].str.contains('MISSING|Missing', case=False, na=False)].to_excel(writer, sheet_name='Missing Items', index=False)
        results_df[results_df['Status'] == '⚠ Needs Install'].to_excel(writer, sheet_name='Needs Install', index=False)
        results_df[results_df['Status'].str.contains('NIC', na=False)].to_excel(writer, sheet_name='NIC Items', index=False)
        results_df[results_df['Status'].isin(['✓ Quoted', '⚡ Included'])].to_excel(writer, sheet_name='Quoted Items', index=False)
        
        all_quotes = [q for qs in quotes.values() for q in qs]
        if all_quotes:
            pd.DataFrame(all_quotes).to_excel(writer, sheet_name='Quote Raw Data', index=False)
    output.seek(0)
    return output

# ===== UI =====
st.markdown('<p class="main-header">📊 Drawing vs Quote Analyzer</p>', unsafe_allow_html=True)
st.caption("Compare equipment schedules against vendor quotations | NIC = Not In Contract")

if not PDF_SUPPORT:
    st.warning("⚠️ PDF support unavailable. Install pdfplumber: `pip install pdfplumber`")

tabs = st.tabs(["📁 Upload & Configure", "📊 Dashboard", "🚨 Missing Items", "🔍 Full Analysis", "📋 Supplier Summary", "💾 Export"])

# ===== TAB 1: Upload & Configure =====
with tabs[0]:
    st.subheader("1️⃣ Upload Drawing/Schedule")
    col1, col2 = st.columns([2, 1])
    with col1:
        draw_file = st.file_uploader("Upload drawing schedule (PDF, Excel, CSV)", type=['pdf', 'csv', 'xlsx', 'xls'], key="draw_upload")
        if draw_file and draw_file.name != st.session_state.drawing_filename:
            with st.spinner("Processing drawing..."):
                dfs = parse_uploaded_file(draw_file, 'drawing')
                if dfs and len(dfs) > 0:
                    combined = max(dfs, key=len).reset_index(drop=True)
                    st.session_state.drawing_df = combined
                    st.session_state.drawing_filename = draw_file.name
                    st.session_state.column_mapping = auto_detect_columns(combined, 'drawing')
                    st.session_state.drawing_data = None
                    st.rerun()
                else:
                    st.error("Could not extract data from file")
    with col2:
        if st.session_state.drawing_filename:
            st.success(f"✅ {st.session_state.drawing_filename}")
    
    if st.session_state.drawing_df is not None:
        st.markdown("---")
        st.subheader("2️⃣ Map Drawing Columns")
        df = st.session_state.drawing_df
        col_options = ['-- Not Used --'] + list(df.columns)
        with st.expander("Preview Drawing Data", expanded=False):
            st.dataframe(df.head(20), height=250, use_container_width=True)
        
        c1, c2, c3 = st.columns(3)
        with c1:
            no_col = st.selectbox("Item No. Column *", col_options, index=col_options.index(st.session_state.column_mapping.get('no')) if st.session_state.column_mapping.get('no') in col_options else 0, key="map_no")
            desc_col = st.selectbox("Description Column *", col_options, index=col_options.index(st.session_state.column_mapping.get('description')) if st.session_state.column_mapping.get('description') in col_options else 0, key="map_desc")
        with c2:
            qty_col = st.selectbox("Quantity Column", col_options, index=col_options.index(st.session_state.column_mapping.get('qty')) if st.session_state.column_mapping.get('qty') in col_options else 0, key="map_qty")
            equip_col = st.selectbox("Equipment/Model # Column", col_options, index=col_options.index(st.session_state.column_mapping.get('equip_num')) if st.session_state.column_mapping.get('equip_num') in col_options else 0, key="map_equip")
        with c3:
            st.session_state.use_categories = st.checkbox("Use Category Codes", value=st.session_state.use_categories)
            cat_col = st.selectbox("Category Column", col_options, index=col_options.index(st.session_state.column_mapping.get('category')) if st.session_state.column_mapping.get('category') in col_options else 0, key="map_cat") if st.session_state.use_categories else '-- Not Used --'
        
        if st.button("✅ Apply Drawing Column Mapping", type="primary"):
            mapping = {k: v for k, v in {'no': no_col, 'description': desc_col, 'qty': qty_col, 'equip_num': equip_col, 'category': cat_col}.items() if v != '-- Not Used --'}
            st.session_state.column_mapping = mapping
            items = extract_drawing_data(df, mapping)
            if items:
                st.session_state.drawing_data = items
                st.success(f"✅ Extracted {len(items)} items from drawing")
                st.rerun()
            else:
                st.error("Could not extract data. Check column mapping.")
    
    if st.session_state.drawing_data:
        with st.expander(f"📋 Extracted Drawing Items ({len(st.session_state.drawing_data)} items)", expanded=False):
            st.dataframe(pd.DataFrame(st.session_state.drawing_data), height=300, use_container_width=True)
    
    st.markdown("---")
    st.subheader("3️⃣ Upload & Configure Quotations")
    quote_input_method = st.radio("Quote Input Method:", ["📁 Upload File", "📋 Paste Data"], horizontal=True)
    
    if quote_input_method == "📁 Upload File":
        quote_files = st.file_uploader("Upload quote files (PDF, Excel, CSV)", type=['pdf', 'csv', 'xlsx', 'xls'], accept_multiple_files=True, key="quote_upload")
        if quote_files:
            for qf in quote_files:
                if qf.name not in st.session_state.quote_dfs:
                    with st.spinner(f"Processing {qf.name}..."):
                        dfs = parse_uploaded_file(qf, 'quote')
                        if dfs and len(dfs) > 0:
                            combined_df = pd.concat(dfs, ignore_index=True) if len(dfs) > 1 else dfs[0]
                            st.session_state.quote_dfs[qf.name] = combined_df.reset_index(drop=True)
                            st.session_state.quote_mappings[qf.name] = auto_detect_columns(combined_df, 'quote')
                            st.success(f"✅ Loaded {qf.name}")
                            st.rerun()
                        else:
                            st.error(f"❌ Could not extract data from {qf.name}")
    else:
        st.caption("Paste quote data in CSV format")
        sample = "Item,Qty,Description,Sell,Sell_Total\n1,,NIC,,\n2,1 ea,WALK IN,97980.27,97980.27\n10,2 ea,STAINLESS,2206.08,4412.16\n11-23,,NIC,,"
        with st.expander("📖 Sample format"):
            st.code(sample)
        pasted_data = st.text_area("Paste your quote data:", height=200)
        quote_name = st.text_input("Name for this quote:", value="Pasted_Quote")
        if st.button("📥 Load Pasted Data", type="primary") and pasted_data.strip():
            try:
                paste_df = clean_dataframe_columns(pd.read_csv(io.StringIO(pasted_data)))
                st.session_state.quote_dfs[quote_name] = paste_df
                st.session_state.quote_mappings[quote_name] = auto_detect_columns(paste_df, 'quote')
                st.success(f"✅ Loaded {len(paste_df)} rows")
                st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")
    
    if st.session_state.quote_dfs:
        st.markdown("---")
        st.markdown("### 📝 Configure Quote Column Mappings")
        for filename, qdf in st.session_state.quote_dfs.items():
            with st.expander(f"📄 {filename} ({len(qdf)} rows)", expanded=(filename not in st.session_state.quotes_data)):
                st.dataframe(qdf.head(25), height=220, use_container_width=True)
                st.markdown("**🎯 Select Columns:**")
                q_col_options = ['-- Not Used --'] + list(qdf.columns)
                current_map = st.session_state.quote_mappings.get(filename, {})
                qc1, qc2, qc3 = st.columns(3)
                with qc1:
                    q_no_col = st.selectbox("Item No. *", q_col_options, index=q_col_options.index(current_map.get('no')) if current_map.get('no') in q_col_options else 0, key=f"qno_{filename}")
                    q_desc_col = st.selectbox("Description", q_col_options, index=q_col_options.index(current_map.get('description')) if current_map.get('description') in q_col_options else 0, key=f"qdesc_{filename}")
                with qc2:
                    q_qty_col = st.selectbox("Quantity", q_col_options, index=q_col_options.index(current_map.get('qty')) if current_map.get('qty') in q_col_options else 0, key=f"qqty_{filename}")
                    q_unit_col = st.selectbox("Unit Price", q_col_options, index=q_col_options.index(current_map.get('unit_price')) if current_map.get('unit_price') in q_col_options else 0, key=f"qunit_{filename}")
                with qc3:
                    q_total_col = st.selectbox("Total Price", q_col_options, index=q_col_options.index(current_map.get('total_price')) if current_map.get('total_price') in q_col_options else 0, key=f"qtotal_{filename}")
                
                bc1, bc2 = st.columns(2)
                with bc1:
                    if st.button(f"✅ Apply Mapping", key=f"apply_{filename}", type="primary"):
                        q_mapping = {k: v for k, v in {'no': q_no_col, 'description': q_desc_col, 'qty': q_qty_col, 'unit_price': q_unit_col, 'total_price': q_total_col}.items() if v != '-- Not Used --'}
                        st.session_state.quote_mappings[filename] = q_mapping
                        items = extract_quote_data(qdf, q_mapping, filename)
                        if items:
                            st.session_state.quotes_data[filename] = items
                            nic_count = sum(1 for i in items if i.get('Is_NIC'))
                            total_val = sum(i['Total_Price'] for i in items if not i.get('Is_NIC'))
                            st.success(f"✅ {len(items)} items ({nic_count} NIC) | ${total_val:,.2f}")
                            st.rerun()
                        else:
                            st.error("No items extracted.")
                with bc2:
                    if st.button(f"🗑️ Remove", key=f"remove_{filename}"):
                        del st.session_state.quote_dfs[filename]
                        st.session_state.quotes_data.pop(filename, None)
                        st.session_state.quote_mappings.pop(filename, None)
                        st.rerun()
                
                if filename in st.session_state.quotes_data:
                    items = st.session_state.quotes_data[filename]
                    st.success(f"✅ {len(items)} items | {sum(1 for i in items if i.get('Is_NIC'))} NIC | ${sum(i['Total_Price'] for i in items if not i.get('Is_NIC')):,.2f}")
                    with st.expander("👁️ View Extracted Items"):
                        st.dataframe(pd.DataFrame(items), height=200, use_container_width=True)
    
    if st.session_state.use_categories:
        st.markdown("---")
        st.subheader("📋 Supplier Code Reference")
        ref_df = pd.DataFrame([{"Code": k, "Description": v, "Quote Required": "No" if k in [1,2,3,8] else "Yes"} for k, v in st.session_state.supplier_codes.items()])
        st.dataframe(ref_df, use_container_width=True, hide_index=True)
    
    st.markdown("---")
    if st.button("🔄 Reset Everything"):
        for key in ['drawing_data', 'drawing_df', 'drawing_filename']:
            st.session_state[key] = None
        for key in ['quotes_data', 'quote_dfs', 'quote_mappings', 'column_mapping']:
            st.session_state[key] = {}
        st.rerun()

# ===== TAB 2: Dashboard =====
with tabs[1]:
    if not st.session_state.drawing_data:
        st.warning("⚠️ Please upload and configure drawing first (Tab 1)")
    elif not st.session_state.quotes_data:
        st.warning("⚠️ Please upload and configure quotations (Tab 1)")
    else:
        all_quotes = [q for qs in st.session_state.quotes_data.values() for q in qs]
        results_df = analyze_data(st.session_state.drawing_data, all_quotes, st.session_state.use_categories, st.session_state.supplier_codes)
        
        missing_critical = len(results_df[results_df['Status'] == '❌ MISSING'])
        if missing_critical > 0:
            st.error(f"🚨 **ALERT: {missing_critical} critical items (Contractor Supply) are NOT in the quote!**")
        
        st.subheader("📊 Coverage Summary")
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("Total Items", len(results_df))
        c2.metric("✓ Quoted", len(results_df[results_df['Status'] == '✓ Quoted']))
        c3.metric("⚡ Included", len(results_df[results_df['Status'] == '⚡ Included']))
        c4.metric("❌ MISSING", missing_critical)
        c5.metric("🚫 NIC", len(results_df[results_df['Status'] == '🚫 NIC']))
        c6.metric("⚠ Needs Install", len(results_df[results_df['Status'] == '⚠ Needs Install']))
        
        col1, col2 = st.columns(2)
        col1.metric("💰 Total Quoted Value", f"${results_df['Total_Price'].sum():,.2f}")
        col2.metric("📦 Items Needing Action", missing_critical + len(results_df[results_df['Status'].isin(['⚠ Needs Install', '⚠ Qty Mismatch'])]))
        
        st.markdown("---")
        ch1, ch2 = st.columns(2)
        with ch1:
            st.subheader("📈 Status Distribution")
            vc = results_df['Status'].value_counts().reset_index()
            vc.columns = ['Status', 'Count']
            colors = {'✓ Quoted': '#28a745', '⚡ Included': '#17a2b8', '❌ MISSING': '#dc3545', '❌ Missing': '#e74c3c', '🚫 NIC': '#6f42c1', '⚠ Qty Mismatch': '#ffc107', '⚠ Needs Install': '#fd7e14', 'Owner Supply': '#6c757d', 'Existing': '#adb5bd', 'N/A': '#e9ecef'}
            fig = px.pie(vc, values='Count', names='Status', color='Status', color_discrete_map=colors, hole=0.4)
            fig.update_layout(height=350)
            st.plotly_chart(fig, use_container_width=True)
        
        with ch2:
            st.subheader("📊 Items by Supplier Code")
            if st.session_state.use_categories:
                code_counts = results_df[results_df['Category'].notna()].groupby('Category').size().reset_index(name='Count')
                code_counts['Category'] = code_counts['Category'].astype(int)
                fig2 = px.bar(code_counts, x='Category', y='Count', color='Category', text='Count')
                fig2.update_layout(height=350, showlegend=False)
                fig2.update_traces(textposition='outside')
                st.plotly_chart(fig2, use_container_width=True)
            else:
                st.info("Enable Category Codes in Tab 1 to see this chart")

# ===== TAB 3: Missing Items =====
with tabs[2]:
    if st.session_state.drawing_data and st.session_state.quotes_data:
        all_quotes = [q for qs in st.session_state.quotes_data.values() for q in qs]
        results_df = analyze_data(st.session_state.drawing_data, all_quotes, st.session_state.use_categories, st.session_state.supplier_codes)
        
        st.subheader("❌ CRITICAL MISSING - Contractor Supply Items")
        st.markdown("*These items require contractor supply but are NOT in the quote:*")
        critical = results_df[results_df['Status'] == '❌ MISSING']
        if len(critical) > 0:
            st.error(f"🚨 {len(critical)} critical items missing!")
            st.dataframe(critical[['Drawing_No', 'Description', 'Drawing_Qty', 'Category', 'Category_Desc', 'Issue']], use_container_width=True, hide_index=True)
        else:
            st.success("✅ All contractor supply items are quoted!")
        
        st.markdown("---")
        st.subheader("⚠ NEEDS INSTALL QUOTE - Owner Supply Items")
        st.markdown("*Owner supplies these items, but contractor needs to quote installation:*")
        needs_install = results_df[results_df['Status'] == '⚠ Needs Install']
        if len(needs_install) > 0:
            st.warning(f"⚠ {len(needs_install)} items need installation quotes")
            st.dataframe(needs_install[['Drawing_No', 'Description', 'Drawing_Qty', 'Category_Desc', 'Issue']], use_container_width=True, hide_index=True)
        else:
            st.success("✅ All installation quotes received!")
        
        st.markdown("---")
        st.subheader("⚠ QTY MISMATCH - Verify with Vendor")
        mismatch = results_df[results_df['Status'] == '⚠ Qty Mismatch']
        if len(mismatch) > 0:
            st.warning(f"⚠ {len(mismatch)} items have quantity mismatches")
            st.dataframe(mismatch[['Drawing_No', 'Description', 'Drawing_Qty', 'Quote_Qty', 'Issue']], use_container_width=True, hide_index=True)
        else:
            st.success("✅ All quantities match!")
        
        st.markdown("---")
        st.subheader("🚫 NIC Items (Not In Contract)")
        nic = results_df[results_df['Status'] == '🚫 NIC']
        if len(nic) > 0:
            st.info(f"{len(nic)} items marked as NIC - excluded from vendor scope")
            st.dataframe(nic[['Drawing_No', 'Description', 'Drawing_Qty']], use_container_width=True, hide_index=True)
    else:
        st.warning("⚠️ Please upload and configure both drawing and quotations first")

# ===== TAB 4: Full Analysis =====
with tabs[3]:
    if st.session_state.drawing_data and st.session_state.quotes_data:
        all_quotes = [q for qs in st.session_state.quotes_data.values() for q in qs]
        results_df = analyze_data(st.session_state.drawing_data, all_quotes, st.session_state.use_categories, st.session_state.supplier_codes)
        
        st.subheader("🔍 Detailed Quote vs Schedule Comparison")
        col1, col2, col3, col4, col5 = st.columns(5)
        show_quoted = col1.checkbox("✓ Quoted", True)
        show_included = col2.checkbox("⚡ Included", True)
        show_missing = col3.checkbox("❌ Missing", True)
        show_nic = col4.checkbox("🚫 NIC", True)
        show_owner = col5.checkbox("Owner/Existing", False)
        
        status_filter = []
        if show_quoted: status_filter.append('✓ Quoted')
        if show_included: status_filter.append('⚡ Included')
        if show_missing: status_filter.extend(['❌ MISSING', '❌ Missing', '⚠ Needs Install', '⚠ Qty Mismatch'])
        if show_nic: status_filter.append('🚫 NIC')
        if show_owner: status_filter.extend(['Owner Supply', 'Existing', 'N/A'])
        
        filtered = results_df[results_df['Status'].isin(status_filter)]
        
        if st.session_state.use_categories:
            codes = st.multiselect("Filter by Supplier Code", list(range(1,9)), default=[4,5,6,7], format_func=lambda x: f"{x}: {st.session_state.supplier_codes[x][:30]}...")
            if codes:
                filtered = filtered[filtered['Category'].isin(codes)]
        
        st.markdown(f"### Results ({len(filtered)} items)")
        
        def color_rows(row):
            colors = {'✓ Quoted': '#d4edda', '⚡ Included': '#d1ecf1', '❌ MISSING': '#f5c6cb', '❌ Missing': '#f8d7da', '🚫 NIC': '#e2d5f0', '⚠ Qty Mismatch': '#fff3cd', '⚠ Needs Install': '#ffe5d0', 'Owner Supply': '#e2e3e5', 'Existing': '#e2e3e5'}
            return [f'background-color: {colors.get(row["Status"], "")}'] * len(row)
        
        display_cols = ['Drawing_No', 'Description', 'Drawing_Qty']
        if st.session_state.use_categories:
            display_cols.extend(['Category', 'Category_Desc'])
        display_cols.extend(['Quote_Item_No', 'Quote_Qty', 'Unit_Price', 'Total_Price', 'Status', 'Issue'])
        
        st.dataframe(filtered[display_cols].style.apply(color_rows, axis=1), use_container_width=True, hide_index=True, height=500)
    else:
        st.warning("⚠️ Please upload and configure both drawing and quotations first")

# ===== TAB 5: Supplier Summary =====
with tabs[4]:
    if st.session_state.drawing_data and st.session_state.quotes_data and st.session_state.use_categories:
        all_quotes = [q for qs in st.session_state.quotes_data.values() for q in qs]
        results_df = analyze_data(st.session_state.drawing_data, all_quotes, st.session_state.use_categories, st.session_state.supplier_codes)
        supplier_summary_df = get_supplier_code_summary(st.session_state.drawing_data, results_df, st.session_state.supplier_codes)
        
        st.subheader("🔢 Supplier Code Summary")
        
        disp_sum = supplier_summary_df.copy()
        disp_sum['Quoted Value'] = disp_sum['Quoted Value'].apply(lambda x: f"${x:,.2f}")
        
        def color_sum(row):
            if row['Quote Required'] == 'No': return ['background-color: #e2e3e5'] * len(row)
            elif row['Missing'] > 0: return ['background-color: #f8d7da'] * len(row)
            elif row['Needs Install'] > 0: return ['background-color: #ffe5d0'] * len(row)
            elif row['Mismatch'] > 0: return ['background-color: #fff3cd'] * len(row)
            elif row['Quoted'] > 0: return ['background-color: #d4edda'] * len(row)
            return [''] * len(row)
        
        st.dataframe(disp_sum.style.apply(color_sum, axis=1), use_container_width=True, hide_index=True)
        
        st.markdown("---")
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("📈 Schedule by Code")
            fig1 = px.bar(supplier_summary_df, x='Code', y='Schedule Items', color='Quote Required', text='Schedule Items')
            fig1.update_traces(textposition='outside')
            fig1.update_layout(height=400)
            st.plotly_chart(fig1, use_container_width=True)
        
        with col2:
            st.subheader("📊 Quote Coverage")
            cov = supplier_summary_df[supplier_summary_df['Quote Required'] == 'Yes']
            fig2 = go.Figure()
            fig2.add_trace(go.Bar(name='Quoted', x=cov['Code'], y=cov['Quoted'], marker_color='#28a745'))
            fig2.add_trace(go.Bar(name='Missing', x=cov['Code'], y=cov['Missing'], marker_color='#dc3545'))
            fig2.add_trace(go.Bar(name='NIC', x=cov['Code'], y=cov['NIC'], marker_color='#6f42c1'))
            fig2.add_trace(go.Bar(name='Needs Install', x=cov['Code'], y=cov['Needs Install'], marker_color='#fd7e14'))
            fig2.update_layout(barmode='stack', height=400)
            st.plotly_chart(fig2, use_container_width=True)
        
        st.markdown("---")
        st.subheader("📋 Items by Supplier Code")
        for code in range(1, 9):
            items = results_df[results_df['Category'] == code]
            if len(items) > 0:
                missing = len(items[items['Status'].str.contains('MISSING|Missing|Needs Install', case=False, na=False)])
                icon = "🚨" if missing > 0 else "✅"
                with st.expander(f"{icon} Code {code}: {st.session_state.supplier_codes[code]} ({len(items)} items)"):
                    st.dataframe(items[['Drawing_No', 'Description', 'Drawing_Qty', 'Quote_Qty', 'Status', 'Issue']], use_container_width=True, hide_index=True)
    elif not st.session_state.use_categories:
        st.warning("⚠️ Enable 'Use Category Codes' in Tab 1 to see Supplier Summary")
    else:
        st.warning("⚠️ Please upload and configure both drawing and quotations first")

# ===== TAB 6: Export =====
with tabs[5]:
    st.subheader("💾 Export Data")
    if st.session_state.drawing_data and st.session_state.quotes_data:
        all_quotes = [q for qs in st.session_state.quotes_data.values() for q in qs]
        results_df = analyze_data(st.session_state.drawing_data, all_quotes, st.session_state.use_categories, st.session_state.supplier_codes)
        supplier_summary_df = get_supplier_code_summary(st.session_state.drawing_data, results_df, st.session_state.supplier_codes) if st.session_state.use_categories else pd.DataFrame()
        
        st.markdown("**Excel Report includes:** Executive Summary | Supplier Code Summary | Full Analysis | Missing Items | NIC Items | Quoted Items | Quote Raw Data")
        
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Preview")
            quoted = len(results_df[results_df['Status'] == '✓ Quoted'])
            included = len(results_df[results_df['Status'] == '⚡ Included'])
            missing = len(results_df[results_df['Status'].str.contains('MISSING|Missing', case=False, na=False)])
            nic = len(results_df[results_df['Status'].str.contains('NIC', na=False)])
            st.dataframe(pd.DataFrame({"Metric": ["Total", "✓ Quoted", "⚡ Included", "❌ Missing", "🚫 NIC", "Value"], "Value": [len(results_df), quoted, included, missing, nic, f"${results_df['Total_Price'].sum():,.2f}"]}), use_container_width=True, hide_index=True)
        
        with col2:
            if st.session_state.use_categories and not supplier_summary_df.empty:
                st.subheader("Supplier Summary")
                st.dataframe(supplier_summary_df[['Code', 'Schedule Items', 'Quoted', 'Missing', 'NIC']], use_container_width=True, hide_index=True)
        
        st.markdown("---")
        col1, col2, col3 = st.columns(3)
        with col1:
            excel = create_excel_report(st.session_state.drawing_data, results_df, st.session_state.quotes_data, supplier_summary_df)
            st.download_button("📥 Full Excel Report", excel, f"Quote_Analysis_{datetime.now().strftime('%Y%m%d')}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
        with col2:
            st.download_button("📥 Analysis CSV", results_df.to_csv(index=False), "Quote_Analysis.csv", "text/csv", use_container_width=True)
        with col3:
            missing_df = results_df[results_df['Status'].str.contains('MISSING|Missing', case=False, na=False)]
            st.download_button(f"📥 Missing Items ({len(missing_df)})", missing_df.to_csv(index=False), "Missing_Items.csv", "text/csv", use_container_width=True)
    else:
        st.warning("⚠️ Please upload and configure both drawing and quotations first")

st.markdown("---")
st.caption("Universal Drawing Quote Analyzer v13.0 | NIC = Not In Contract")
