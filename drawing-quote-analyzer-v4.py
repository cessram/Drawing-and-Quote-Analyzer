import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re
from datetime import datetime

try:
    import pdfplumber
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False

st.set_page_config(page_title="Drawing Quote Analyzer", page_icon="📊", layout="wide")

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
if 'drawing_data' not in st.session_state:
    st.session_state.drawing_data = None
if 'drawing_df' not in st.session_state:
    st.session_state.drawing_df = None
if 'quotes_data' not in st.session_state:
    st.session_state.quotes_data = {}
if 'quote_dfs' not in st.session_state:
    st.session_state.quote_dfs = {}
if 'quote_mappings' not in st.session_state:
    st.session_state.quote_mappings = {}
if 'drawing_filename' not in st.session_state:
    st.session_state.drawing_filename = None
if 'column_mapping' not in st.session_state:
    st.session_state.column_mapping = {}
if 'supplier_codes' not in st.session_state:
    st.session_state.supplier_codes = DEFAULT_SUPPLIER_CODES.copy()
if 'use_categories' not in st.session_state:
    st.session_state.use_categories = True

def clean_dataframe_columns(df):
    """Clean DataFrame column names to ensure uniqueness."""
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
    """Extract quote data from PDF using intelligent text parsing."""
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
                    
                    # Skip header/footer lines
                    skip_patterns = [
                        'Canadian Restaurant Supply', 'Bird Construc', 'Page ', 
                        'FWG LTC', 'Quote valid', 'Inspections:', 'Item Qty Description',
                        'ITEM TOTAL:', 'Merchandise', 'GST', 'Tax', 'Total'
                    ]
                    if any(p in line for p in skip_patterns):
                        continue
                    
                    # Pattern 1: Item range with NIC (e.g., "11-23 NIC" or "52-66 NIC")
                    range_nic = re.match(r'^(\d+)[-–](\d+)\s+NIC\s*$', line, re.IGNORECASE)
                    if range_nic:
                        start, end = int(range_nic.group(1)), int(range_nic.group(2))
                        for num in range(start, end + 1):
                            items.append({'Item': str(num), 'Qty': '', 'Description': 'NIC', 'Sell': '', 'Sell_Total': ''})
                        continue
                    
                    # Pattern 2: Single item NIC (e.g., "1 NIC" or "44 NIC")
                    single_nic = re.match(r'^(\d+)\s+NIC\s*$', line, re.IGNORECASE)
                    if single_nic:
                        items.append({'Item': single_nic.group(1), 'Qty': '', 'Description': 'NIC', 'Sell': '', 'Sell_Total': ''})
                        continue
                    
                    # Pattern 3: Item with qty, description, and prices
                    # e.g., "2 1 ea WALK IN $97,980.27 $97,980.27"
                    # e.g., "10 2 ea STAINLESS $2,206.08 $4,412.16"
                    item_full = re.match(
                        r'^(\d+)\s+(\d+)\s*ea\s+([A-Z][A-Z0-9\s,./\-&\(\)\'\"]+?)\s+\$?([\d,]+\.?\d*)\s+\$?([\d,]+\.?\d*)\s*$',
                        line, re.IGNORECASE
                    )
                    if item_full:
                        items.append({
                            'Item': item_full.group(1),
                            'Qty': f"{item_full.group(2)} ea",
                            'Description': item_full.group(3).strip(),
                            'Sell': item_full.group(4),
                            'Sell_Total': item_full.group(5)
                        })
                        continue
                    
                    # Pattern 4: Item with description and single price (unit = total)
                    item_single_price = re.match(
                        r'^(\d+)\s+(\d+)\s*ea\s+([A-Z][A-Z0-9\s,./\-&\(\)\'\"]+?)\s+\$?([\d,]+\.?\d*)\s*$',
                        line, re.IGNORECASE
                    )
                    if item_single_price:
                        price = item_single_price.group(4)
                        items.append({
                            'Item': item_single_price.group(1),
                            'Qty': f"{item_single_price.group(2)} ea",
                            'Description': item_single_price.group(3).strip(),
                            'Sell': price,
                            'Sell_Total': price
                        })
                        continue
                    
                    # Pattern 5: Item number at start with description containing price at end
                    item_desc_price = re.match(
                        r'^(\d+)\s+(\d+)\s*ea\s+(.+?)\s+\$?([\d,]+\.?\d*)\s+\$?([\d,]+\.?\d*)\s*$',
                        line
                    )
                    if item_desc_price:
                        items.append({
                            'Item': item_desc_price.group(1),
                            'Qty': f"{item_desc_price.group(2)} ea",
                            'Description': item_desc_price.group(3).strip(),
                            'Sell': item_desc_price.group(4),
                            'Sell_Total': item_desc_price.group(5)
                        })
                        continue
        
        if items:
            return pd.DataFrame(items)
    except Exception as e:
        st.warning(f"Text extraction error: {e}")
    
    return None

def parse_pdf_tables_for_quote(uploaded_file):
    """Extract tables from PDF with improved handling for quote format."""
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
                    
                    # Find header row
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
                                row_dict = {}
                                for j, cell in enumerate(row):
                                    if j < len(headers):
                                        row_dict[headers[j]] = cell
                                all_rows.append(row_dict)
                    else:
                        # No header - use positional columns
                        for row in table:
                            if row and len(row) >= 2:
                                first_cell = str(row[0]).strip() if row[0] else ''
                                # Check if first cell looks like item number
                                if re.match(r'^\d+(-\d+)?$', first_cell):
                                    all_rows.append({
                                        'Item': row[0],
                                        'Qty': row[1] if len(row) > 1 else '',
                                        'Description': row[2] if len(row) > 2 else '',
                                        'Sell': row[3] if len(row) > 3 else '',
                                        'Sell Total': row[4] if len(row) > 4 else ''
                                    })
        
        if all_rows:
            df = pd.DataFrame(all_rows)
            df = clean_dataframe_columns(df)
            return [df]
    except Exception as e:
        st.warning(f"PDF table extraction error: {e}")
    
    return None

def parse_pdf_tables(uploaded_file):
    """Extract tables from PDF - for drawings."""
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
    """Parse Excel file to list of DataFrames."""
    try:
        uploaded_file.seek(0)
        xl = pd.ExcelFile(uploaded_file)
        dfs = []
        for name in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name=name)
            df = clean_dataframe_columns(df)
            if len(df) > 0:
                dfs.append(df)
        return dfs
    except Exception as e:
        st.warning(f"Could not read Excel file: {e}")
        return None

def parse_csv_file(uploaded_file):
    """Parse CSV file to DataFrame."""
    try:
        uploaded_file.seek(0)
        df = pd.read_csv(uploaded_file)
        df = clean_dataframe_columns(df)
        return [df] if len(df) > 0 else None
    except Exception as e:
        st.warning(f"Could not read CSV file: {e}")
        return None

def parse_uploaded_file(uploaded_file, file_type='drawing'):
    """Parse any supported file type."""
    ext = uploaded_file.name.split('.')[-1].lower()
    
    if ext == 'pdf':
        if file_type == 'quote':
            # Try text extraction first for quotes
            text_df = extract_quote_from_pdf_text(uploaded_file)
            if text_df is not None and len(text_df) > 0:
                return [text_df]
            # Fall back to table extraction
            return parse_pdf_tables_for_quote(uploaded_file)
        return parse_pdf_tables(uploaded_file)
    elif ext in ['xlsx', 'xls']:
        return parse_excel_file(uploaded_file)
    elif ext == 'csv':
        return parse_csv_file(uploaded_file)
    return None

def auto_detect_columns(df, file_type='drawing'):
    """Auto-detect column mappings."""
    cols_lower = {c: c.lower().strip() for c in df.columns}
    
    if file_type == 'drawing':
        patterns = {
            'no': ['no', 'no.', 'item', 'item #', 'item no', 'number', '#', 'id'],
            'description': ['description', 'desc', 'equipment', 'name', 'material'],
            'qty': ['qty', 'qty.', 'quantity', 'count', 'amount'],
            'category': ['category', 'cat', 'supplier code', 'code', 'type'],
            'equip_num': ['equipment number', 'equip num', 'equip no', 'model', 'part no'],
        }
    else:
        patterns = {
            'no': ['item', 'no', 'no.', 'item #', 'number', '#', 'id', 'line'],
            'description': ['description', 'desc', 'equipment', 'name', 'material', 'product'],
            'qty': ['qty', 'qty.', 'quantity', 'count', 'ea'],
            'unit_price': ['sell', 'unit price', 'price', 'rate', 'unit cost', 'each', 'unit'],
            'total_price': ['sell_total', 'sell total', 'total', 'total price', 'ext price', 'extended', 'amount'],
        }
    
    found = {}
    for key, opts in patterns.items():
        for col, col_low in cols_lower.items():
            if col_low in opts:
                found[key] = col
                break
            for opt in opts:
                if opt in col_low:
                    found[key] = col
                    break
            if key in found:
                break
    return found

def clean_numeric(val):
    """Clean numeric value from string."""
    if pd.isna(val):
        return None
    val_str = str(val).strip()
    val_str = re.sub(r'[,$]', '', val_str)
    val_str = re.sub(r'[^\d.\-]', '', val_str)
    try:
        return float(val_str) if val_str else None
    except:
        return None

def parse_qty_value(val):
    """Parse quantity from formats like '1 ea', '2 ea', '1', etc."""
    if pd.isna(val):
        return 1
    val_str = str(val).strip().lower()
    if not val_str or val_str in ('nan', 'none', ''):
        return 1
    match = re.search(r'(\d+)\s*ea', val_str)
    if match:
        return int(match.group(1))
    num = clean_numeric(val)
    if num and num > 0:
        return int(num)
    return 1

def extract_drawing_data(df, col_map):
    """Extract drawing data using column mapping."""
    items = []
    no_col = col_map.get('no')
    desc_col = col_map.get('description')
    qty_col = col_map.get('qty')
    cat_col = col_map.get('category')
    equip_col = col_map.get('equip_num')
    
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
            
            qty = 1
            if qty_col:
                qty_val = clean_numeric(row.get(qty_col, 1))
                if qty_val:
                    qty = int(qty_val)
            
            cat = None
            if cat_col:
                cat_val = clean_numeric(row.get(cat_col, ''))
                if cat_val:
                    cat = int(cat_val)
            
            equip_num = '-'
            if equip_col:
                en = str(row.get(equip_col, '')).strip()
                if en and en.lower() not in ('nan', '', '-', 'none'):
                    equip_num = en
            
            items.append({
                'No': no_val,
                'Equip_Num': equip_num,
                'Description': desc_val,
                'Qty': qty,
                'Category': cat
            })
        except:
            continue
    
    return items if items else None

def extract_quote_data(df, col_map, source_file):
    """Extract quote data using column mapping."""
    items = []
    
    no_col = col_map.get('no')
    desc_col = col_map.get('description')
    qty_col = col_map.get('qty')
    unit_col = col_map.get('unit_price')
    total_col = col_map.get('total_price')
    
    for idx, row in df.iterrows():
        try:
            # Get item number
            no_val = ''
            if no_col and no_col in df.columns:
                no_val = str(row.get(no_col, '')).strip()
                if no_val.lower() in ('nan', 'none', 'item', ''):
                    no_val = ''
            
            if not no_val:
                continue
            
            # Get description
            desc_val = ''
            if desc_col and desc_col in df.columns:
                desc_val = str(row.get(desc_col, '')).strip()
                if desc_val.lower() in ('nan', 'none', 'description'):
                    desc_val = ''
            
            # Check for NIC
            is_nic = False
            if desc_val.upper().strip() == 'NIC' or 'NIC' in desc_val.upper():
                is_nic = True
            
            # Check qty column for NIC too
            qty_raw = ''
            if qty_col and qty_col in df.columns:
                qty_raw = str(row.get(qty_col, '')).strip()
                if qty_raw.upper() == 'NIC':
                    is_nic = True
                    desc_val = 'NIC'
            
            # Get quantity
            qty = 1
            if qty_col and qty_col in df.columns and not is_nic:
                qty = parse_qty_value(row.get(qty_col, ''))
            
            # Get prices
            unit_price = 0.0
            if unit_col and unit_col in df.columns:
                up = clean_numeric(row.get(unit_col, 0))
                if up:
                    unit_price = up
            
            total_price = 0.0
            if total_col and total_col in df.columns:
                tp = clean_numeric(row.get(total_col, 0))
                if tp:
                    total_price = tp
            
            if total_price == 0 and unit_price > 0:
                total_price = unit_price * qty
            if unit_price == 0 and total_price > 0 and qty > 0:
                unit_price = total_price / qty
            
            # Handle item ranges like "11-23"
            range_match = re.match(r'^(\d+)[-–](\d+)$', no_val)
            if range_match:
                start, end = int(range_match.group(1)), int(range_match.group(2))
                for num in range(start, end + 1):
                    items.append({
                        'Item_No': str(num),
                        'Description': desc_val if desc_val else 'NIC',
                        'Qty': qty,
                        'Unit_Price': 0,
                        'Total_Price': 0,
                        'Is_NIC': True,
                        'Source_File': source_file
                    })
            else:
                items.append({
                    'Item_No': no_val,
                    'Description': desc_val if desc_val else '-',
                    'Qty': qty,
                    'Unit_Price': unit_price,
                    'Total_Price': total_price,
                    'Is_NIC': is_nic,
                    'Source_File': source_file
                })
        except:
            continue
    
    return items

def match_items(drawing_no, quotes):
    """Match drawing item to quote by item number."""
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
    """Analyze drawing items against quotes."""
    if supplier_codes is None:
        supplier_codes = DEFAULT_SUPPLIER_CODES
    
    analysis = []
    for item in drawing_items:
        match = match_items(item['No'], quotes)
        cat = item.get('Category')
        
        if use_categories and cat in [1, 2, 3]:
            status, issue = "Owner Supply", supplier_codes.get(cat, "Owner handles")
        elif use_categories and cat == 8:
            status, issue = "Existing", "Existing or relocated"
        elif item.get('Description', '').upper() in ('SPARE', '-', 'N/A'):
            status, issue = "N/A", "Spare or placeholder"
        elif match:
            if match.get('Is_NIC'):
                status, issue = "NIC", "Not In Contract"
            elif match['Qty'] == item['Qty']:
                status, issue = "Quoted", None
            else:
                status, issue = "Qty Mismatch", f"Drawing: {item['Qty']}, Quote: {match['Qty']}"
        else:
            if use_categories and cat == 7:
                status, issue = "Needs Pricing", "Owner supplies - needs install pricing"
            elif use_categories and cat in [5, 6]:
                status, issue = "MISSING", "Critical - requires quote"
            else:
                status, issue = "MISSING", "Not found in quotes"
        
        analysis.append({
            'Drawing_No': item['No'],
            'Equip_Num': item.get('Equip_Num', '-'),
            'Description': item['Description'],
            'Drawing_Qty': item['Qty'],
            'Category': cat,
            'Category_Desc': supplier_codes.get(cat, '-') if cat and use_categories else '-',
            'Quote_Item_No': match['Item_No'] if match else '-',
            'Quote_Qty': match['Qty'] if match else 0,
            'Unit_Price': match['Unit_Price'] if match else 0,
            'Total_Price': match['Total_Price'] if match and not match.get('Is_NIC') else 0,
            'Quote_Source': match['Source_File'] if match else '-',
            'Status': status,
            'Issue': issue
        })
    
    return pd.DataFrame(analysis)

# ===== UI =====
st.markdown("## 📊 Drawing vs Quote Analyzer")
st.caption("Compare equipment schedules against vendor quotations | NIC = Not In Contract")

if not PDF_SUPPORT:
    st.warning("⚠️ PDF support unavailable. Install pdfplumber: `pip install pdfplumber`")

tabs = st.tabs(["📁 Upload & Configure", "📊 Dashboard", "🔍 Analysis", "📋 Summary", "💾 Export"])

# ===== TAB 1: Upload & Configure =====
with tabs[0]:
    st.subheader("1️⃣ Upload Drawing/Schedule")
    
    col1, col2 = st.columns([2, 1])
    with col1:
        draw_file = st.file_uploader("Upload drawing schedule (PDF, Excel, CSV)", type=['pdf', 'csv', 'xlsx', 'xls'], key="draw_upload")
        
        if draw_file:
            if draw_file.name != st.session_state.drawing_filename:
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
    
    # Column mapping for drawing
    if st.session_state.drawing_df is not None:
        st.markdown("---")
        st.subheader("2️⃣ Map Drawing Columns")
        
        df = st.session_state.drawing_df
        col_options = ['-- Not Used --'] + list(df.columns)
        
        with st.expander("Preview Drawing Data", expanded=False):
            st.dataframe(df.head(20), height=250, use_container_width=True)
        
        c1, c2, c3 = st.columns(3)
        with c1:
            no_idx = col_options.index(st.session_state.column_mapping.get('no')) if st.session_state.column_mapping.get('no') in col_options else 0
            no_col = st.selectbox("Item No. Column *", col_options, index=no_idx, key="map_no")
            desc_idx = col_options.index(st.session_state.column_mapping.get('description')) if st.session_state.column_mapping.get('description') in col_options else 0
            desc_col = st.selectbox("Description Column *", col_options, index=desc_idx, key="map_desc")
        
        with c2:
            qty_idx = col_options.index(st.session_state.column_mapping.get('qty')) if st.session_state.column_mapping.get('qty') in col_options else 0
            qty_col = st.selectbox("Quantity Column", col_options, index=qty_idx, key="map_qty")
            equip_idx = col_options.index(st.session_state.column_mapping.get('equip_num')) if st.session_state.column_mapping.get('equip_num') in col_options else 0
            equip_col = st.selectbox("Equipment/Model # Column", col_options, index=equip_idx, key="map_equip")
        
        with c3:
            st.session_state.use_categories = st.checkbox("Use Category Codes", value=st.session_state.use_categories)
            if st.session_state.use_categories:
                cat_idx = col_options.index(st.session_state.column_mapping.get('category')) if st.session_state.column_mapping.get('category') in col_options else 0
                cat_col = st.selectbox("Category Column", col_options, index=cat_idx, key="map_cat")
            else:
                cat_col = '-- Not Used --'
        
        if st.button("✅ Apply Drawing Column Mapping", type="primary"):
            mapping = {}
            if no_col != '-- Not Used --': mapping['no'] = no_col
            if desc_col != '-- Not Used --': mapping['description'] = desc_col
            if qty_col != '-- Not Used --': mapping['qty'] = qty_col
            if equip_col != '-- Not Used --': mapping['equip_num'] = equip_col
            if cat_col != '-- Not Used --': mapping['category'] = cat_col
            
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
    
    # ===== QUOTE SECTION =====
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
                            combined_df = combined_df.reset_index(drop=True)
                            st.session_state.quote_dfs[qf.name] = combined_df
                            st.session_state.quote_mappings[qf.name] = auto_detect_columns(combined_df, 'quote')
                            st.success(f"✅ Loaded {qf.name} ({len(combined_df)} rows)")
                            st.rerun()
                        else:
                            st.error(f"❌ Could not extract data from {qf.name}")
                            st.info("💡 Try 'Paste Data' option - copy from PDF and paste in CSV format")
    
    else:  # Paste Data
        st.markdown("**Paste quote data in CSV format:**")
        st.caption("Copy from Excel/PDF and paste. Format: Item, Qty, Description, Sell, Sell Total")
        
        sample = """Item,Qty,Description,Sell,Sell_Total
1,,NIC,,
2,1 ea,WALK IN,97980.27,97980.27
3,1 ea,WALK IN,70727.09,70727.09
10,2 ea,STAINLESS,2206.08,4412.16
11-23,,NIC,,
24,1 ea,INGREDIENT BIN,386.48,395.40
37,1 ea,HOOD,66746.05,66746.05
42,1 ea,KETTLE ELECTRIC,9111.11,14790.11
47,2 ea,COMBI OVEN,20582.70,50568.76"""
        
        with st.expander("📖 See sample format"):
            st.code(sample)
        
        pasted_data = st.text_area("Paste your quote data:", height=200)
        quote_name = st.text_input("Name for this quote:", value="Pasted_Quote")
        
        if st.button("📥 Load Pasted Data", type="primary"):
            if pasted_data.strip():
                try:
                    paste_df = pd.read_csv(io.StringIO(pasted_data))
                    paste_df = clean_dataframe_columns(paste_df)
                    st.session_state.quote_dfs[quote_name] = paste_df
                    st.session_state.quote_mappings[quote_name] = auto_detect_columns(paste_df, 'quote')
                    st.success(f"✅ Loaded {len(paste_df)} rows")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")
            else:
                st.warning("Please paste data first")
    
    # Configure each quote file with MANUAL COLUMN SELECTION
    if st.session_state.quote_dfs:
        st.markdown("---")
        st.markdown("### 📝 Configure Quote Column Mappings")
        
        for filename, qdf in st.session_state.quote_dfs.items():
            with st.expander(f"📄 {filename} ({len(qdf)} rows)", expanded=(filename not in st.session_state.quotes_data)):
                st.markdown("**Raw Data Preview:**")
                st.dataframe(qdf.head(25), height=220, use_container_width=True)
                
                st.markdown("---")
                st.markdown("**🎯 Select Columns for Analysis:**")
                
                q_col_options = ['-- Not Used --'] + list(qdf.columns)
                current_map = st.session_state.quote_mappings.get(filename, {})
                
                qc1, qc2, qc3 = st.columns(3)
                
                with qc1:
                    q_no_idx = q_col_options.index(current_map.get('no')) if current_map.get('no') in q_col_options else 0
                    q_no_col = st.selectbox("Item No. Column *", q_col_options, index=q_no_idx, key=f"qno_{filename}")
                    
                    q_desc_idx = q_col_options.index(current_map.get('description')) if current_map.get('description') in q_col_options else 0
                    q_desc_col = st.selectbox("Description Column", q_col_options, index=q_desc_idx, key=f"qdesc_{filename}")
                
                with qc2:
                    q_qty_idx = q_col_options.index(current_map.get('qty')) if current_map.get('qty') in q_col_options else 0
                    q_qty_col = st.selectbox("Quantity Column", q_col_options, index=q_qty_idx, key=f"qqty_{filename}")
                    
                    q_unit_idx = q_col_options.index(current_map.get('unit_price')) if current_map.get('unit_price') in q_col_options else 0
                    q_unit_col = st.selectbox("Unit Price (Sell)", q_col_options, index=q_unit_idx, key=f"qunit_{filename}")
                
                with qc3:
                    q_total_idx = q_col_options.index(current_map.get('total_price')) if current_map.get('total_price') in q_col_options else 0
                    q_total_col = st.selectbox("Total Price (Sell Total)", q_col_options, index=q_total_idx, key=f"qtotal_{filename}")
                
                bcol1, bcol2 = st.columns(2)
                
                with bcol1:
                    if st.button(f"✅ Apply Mapping", key=f"apply_{filename}", type="primary"):
                        q_mapping = {}
                        if q_no_col != '-- Not Used --': q_mapping['no'] = q_no_col
                        if q_desc_col != '-- Not Used --': q_mapping['description'] = q_desc_col
                        if q_qty_col != '-- Not Used --': q_mapping['qty'] = q_qty_col
                        if q_unit_col != '-- Not Used --': q_mapping['unit_price'] = q_unit_col
                        if q_total_col != '-- Not Used --': q_mapping['total_price'] = q_total_col
                        
                        st.session_state.quote_mappings[filename] = q_mapping
                        items = extract_quote_data(qdf, q_mapping, filename)
                        
                        if items:
                            st.session_state.quotes_data[filename] = items
                            nic_count = sum(1 for i in items if i.get('Is_NIC'))
                            total_val = sum(i['Total_Price'] for i in items if not i.get('Is_NIC'))
                            st.success(f"✅ Extracted {len(items)} items ({nic_count} NIC) | ${total_val:,.2f}")
                            st.rerun()
                        else:
                            st.error("No items extracted. Check column mapping.")
                
                with bcol2:
                    if st.button(f"🗑️ Remove Quote", key=f"remove_{filename}"):
                        del st.session_state.quote_dfs[filename]
                        if filename in st.session_state.quotes_data:
                            del st.session_state.quotes_data[filename]
                        if filename in st.session_state.quote_mappings:
                            del st.session_state.quote_mappings[filename]
                        st.rerun()
                
                # Show extracted items if available
                if filename in st.session_state.quotes_data:
                    items = st.session_state.quotes_data[filename]
                    nic_count = sum(1 for i in items if i.get('Is_NIC'))
                    total = sum(i['Total_Price'] for i in items if not i.get('Is_NIC'))
                    st.success(f"✅ Processed: {len(items)} items | {nic_count} NIC | Total: ${total:,.2f}")
                    
                    with st.expander("👁️ View Extracted Quote Items"):
                        st.dataframe(pd.DataFrame(items), height=250, use_container_width=True)
    
    # Summary of loaded quotes
    if st.session_state.quotes_data:
        st.markdown("---")
        st.subheader("📊 Loaded Quotes Summary")
        for fn, qs in st.session_state.quotes_data.items():
            nic_count = sum(1 for q in qs if q.get('Is_NIC'))
            total = sum(q['Total_Price'] for q in qs if not q.get('Is_NIC'))
            st.caption(f"• **{fn}**: {len(qs)} items ({nic_count} NIC) = ${total:,.2f}")
    
    st.markdown("---")
    if st.button("🔄 Reset Everything"):
        for key in ['drawing_data', 'drawing_df', 'quotes_data', 'quote_dfs', 'quote_mappings', 'drawing_filename', 'column_mapping']:
            st.session_state[key] = None if 'data' in key or 'df' in key or 'filename' in key else {}
        st.rerun()

# ===== TAB 2: Dashboard =====
with tabs[1]:
    if not st.session_state.drawing_data:
        st.warning("⚠️ Please upload and configure drawing first (Tab 1)")
    elif not st.session_state.quotes_data:
        st.warning("⚠️ Please upload and configure quotations (Tab 1)")
        if st.session_state.quote_dfs:
            st.info("💡 Quote files loaded but not processed. Go to Tab 1 and click 'Apply Mapping'.")
    else:
        all_quotes = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_data(st.session_state.drawing_data, all_quotes, st.session_state.use_categories, st.session_state.supplier_codes)
        
        st.subheader("📊 Coverage Summary")
        
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("✅ Quoted", len(df[df['Status'] == 'Quoted']))
        c2.metric("❌ Missing", len(df[df['Status'] == 'MISSING']))
        c3.metric("⚠️ Qty Mismatch", len(df[df['Status'] == 'Qty Mismatch']))
        c4.metric("🚫 NIC", len(df[df['Status'] == 'NIC']))
        c5.metric("📋 Needs Pricing", len(df[df['Status'] == 'Needs Pricing']))
        
        col1, col2 = st.columns(2)
        col1.metric("💰 Total Quoted Value", f"${df['Total_Price'].sum():,.2f}")
        col2.metric("📦 Total Items", len(df))
        
        ch1, ch2 = st.columns(2)
        with ch1:
            vc = df['Status'].value_counts().reset_index()
            vc.columns = ['Status', 'Count']
            colors = {'Quoted': '#28a745', 'MISSING': '#dc3545', 'Qty Mismatch': '#ffc107', 'NIC': '#6f42c1', 'Needs Pricing': '#fd7e14', 'Owner Supply': '#6c757d', 'Existing': '#adb5bd', 'N/A': '#e9ecef'}
            fig = px.pie(vc, values='Count', names='Status', color='Status', color_discrete_map=colors, title="Status Distribution")
            st.plotly_chart(fig, use_container_width=True)
        
        with ch2:
            fig2 = px.bar(vc, x='Status', y='Count', title="Items by Status", color='Status', color_discrete_map=colors)
            st.plotly_chart(fig2, use_container_width=True)

# ===== TAB 3: Analysis =====
with tabs[2]:
    if st.session_state.drawing_data and st.session_state.quotes_data:
        all_quotes = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_data(st.session_state.drawing_data, all_quotes, st.session_state.use_categories, st.session_state.supplier_codes)
        
        st.subheader("🔍 Detailed Analysis")
        
        status_opts = df['Status'].unique().tolist()
        filt_status = st.multiselect("Filter by Status", status_opts, default=status_opts)
        fdf = df[df['Status'].isin(filt_status)]
        
        def highlight_row(row):
            colors = {'Quoted': 'background-color:#d4edda', 'MISSING': 'background-color:#f8d7da', 'Qty Mismatch': 'background-color:#fff3cd', 'NIC': 'background-color:#e2d5f0', 'Needs Pricing': 'background-color:#ffe5d0'}
            return [colors.get(row['Status'], '')] * len(row)
        
        display_cols = ['Drawing_No', 'Equip_Num', 'Description', 'Drawing_Qty']
        if st.session_state.use_categories:
            display_cols.append('Category')
        display_cols.extend(['Quote_Item_No', 'Quote_Qty', 'Unit_Price', 'Total_Price', 'Status', 'Issue'])
        
        st.dataframe(fdf[display_cols].style.apply(highlight_row, axis=1), height=450, use_container_width=True)
        
        st.subheader("🚨 Critical Missing Items")
        critical = df[df['Status'] == 'MISSING']
        if len(critical) > 0:
            st.error(f"⚠️ {len(critical)} items require quotes!")
            st.dataframe(critical[['Drawing_No', 'Equip_Num', 'Description', 'Drawing_Qty']], use_container_width=True)
        else:
            st.success("✅ No critical missing items!")
        
        nic_items = df[df['Status'] == 'NIC']
        if len(nic_items) > 0:
            st.subheader("🚫 NIC Items (Not In Contract)")
            st.dataframe(nic_items[['Drawing_No', 'Description', 'Drawing_Qty']], use_container_width=True)
    else:
        st.warning("⚠️ Please upload and configure both drawing and quotations first")

# ===== TAB 4: Summary =====
with tabs[3]:
    if st.session_state.drawing_data and st.session_state.quotes_data:
        all_quotes = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_data(st.session_state.drawing_data, all_quotes, st.session_state.use_categories, st.session_state.supplier_codes)
        
        st.subheader("📋 Summary by Status")
        status_summary = df.groupby('Status').agg(Items=('Status', 'count'), Total_Value=('Total_Price', 'sum')).reset_index()
        status_summary['Total_Value'] = status_summary['Total_Value'].apply(lambda x: f"${x:,.2f}")
        st.dataframe(status_summary, use_container_width=True)
        
        st.subheader("📄 Quote Files Summary")
        for fn, qs in st.session_state.quotes_data.items():
            nic_count = sum(1 for q in qs if q.get('Is_NIC'))
            total = sum(q['Total_Price'] for q in qs if not q.get('Is_NIC'))
            st.caption(f"• **{fn}**: {len(qs)} items ({nic_count} NIC) = ${total:,.2f}")
    else:
        st.warning("⚠️ Please upload and configure both drawing and quotations first")

# ===== TAB 5: Export =====
with tabs[4]:
    st.subheader("💾 Export Data")
    
    if st.session_state.drawing_data and st.session_state.quotes_data:
        all_quotes = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_data(st.session_state.drawing_data, all_quotes, st.session_state.use_categories, st.session_state.supplier_codes)
        
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Full Analysis', index=False)
            df[df['Status'] == 'MISSING'].to_excel(writer, sheet_name='Missing Items', index=False)
            df[df['Status'] == 'Quoted'].to_excel(writer, sheet_name='Quoted Items', index=False)
            df[df['Status'] == 'NIC'].to_excel(writer, sheet_name='NIC Items', index=False)
            pd.DataFrame(all_quotes).to_excel(writer, sheet_name='All Quotes Raw', index=False)
        out.seek(0)
        
        st.download_button("📥 Download Full Analysis Report", out, f"Analysis_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", type="primary")
    else:
        st.warning("⚠️ Please upload and configure both drawing and quotations first")

st.markdown("---")
st.caption("Universal Drawing Quote Analyzer v12.0 | NIC = Not In Contract")
