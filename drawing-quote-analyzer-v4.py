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

def parse_pdf_tables(uploaded_file):
    """Extract tables from PDF."""
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

def parse_uploaded_file(uploaded_file):
    """Parse any supported file type and return list of DataFrames."""
    ext = uploaded_file.name.split('.')[-1].lower()
    
    if ext == 'pdf':
        return parse_pdf_tables(uploaded_file)
    elif ext in ['xlsx', 'xls']:
        return parse_excel_file(uploaded_file)
    elif ext == 'csv':
        return parse_csv_file(uploaded_file)
    return None

def auto_detect_columns(df, file_type='drawing'):
    """Auto-detect column mappings based on common naming patterns."""
    cols_lower = {c: c.lower().strip() for c in df.columns}
    
    if file_type == 'drawing':
        patterns = {
            'no': ['no', 'no.', 'item', 'item #', 'item no', 'number', '#', 'id', 'ref'],
            'description': ['description', 'desc', 'equipment', 'name', 'item description', 'material'],
            'qty': ['qty', 'qty.', 'quantity', 'count', 'amount', 'units'],
            'category': ['category', 'cat', 'supplier code', 'code', 'type', 'supply'],
            'equip_num': ['equipment number', 'equip num', 'equip no', 'equip #', 'model', 'part no', 'part #', 'new equipment'],
            'unit': ['unit', 'uom', 'measure'],
            'remarks': ['remarks', 'notes', 'comment', 'comments']
        }
    else:  # quote
        patterns = {
            'no': ['item', 'no', 'no.', 'item #', 'item no', 'number', '#', 'id', 'ref', 'line', 'seq'],
            'description': ['description', 'desc', 'equipment', 'name', 'item description', 'material', 'product', 'model'],
            'qty': ['qty', 'qty.', 'quantity', 'count', 'amount', 'units', 'ea'],
            'unit_price': ['sell', 'unit price', 'price', 'rate', 'unit cost', 'cost ea', 'each', 'unit'],
            'total_price': ['sell total', 'total', 'total price', 'ext price', 'extended', 'amount', 'line total', 'ext'],
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
    # Handle "X ea" format
    match = re.search(r'(\d+)\s*ea', val_str)
    if match:
        return int(match.group(1))
    # Handle plain numbers
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
            
            # Skip header rows and empty rows
            if not no_val or no_val.lower() in ('nan', '', 'no', 'no.', 'item', 'none'):
                continue
            if not desc_val or desc_val.lower() in ('nan', '', 'description', 'none'):
                continue
            
            # Get quantity
            qty = 1
            if qty_col:
                qty_val = clean_numeric(row.get(qty_col, 1))
                if qty_val:
                    qty = int(qty_val)
            
            # Get category
            cat = None
            if cat_col:
                cat_val = clean_numeric(row.get(cat_col, ''))
                if cat_val:
                    cat = int(cat_val)
            
            # Get equipment number
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
    """Extract quote data using column mapping with NIC handling."""
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
            if no_col:
                no_val = str(row.get(no_col, '')).strip()
                if no_val.lower() in ('nan', 'none'):
                    no_val = ''
            
            # Get description
            desc_val = ''
            if desc_col:
                desc_val = str(row.get(desc_col, '')).strip()
                if desc_val.lower() in ('nan', 'none'):
                    desc_val = ''
            
            # Skip if both are empty or look like headers
            if not no_val and not desc_val:
                continue
            if no_val.lower() in ('item', 'no', 'no.', '#', 'line'):
                continue
            if desc_val.lower() in ('description', 'item description'):
                continue
            
            # Check for NIC (Not In Contract)
            is_nic = desc_val.upper().strip() == 'NIC' or 'NIC' in desc_val.upper()
            
            # Get quantity
            qty = 1
            if qty_col:
                qty_raw = row.get(qty_col, '')
                if pd.notna(qty_raw) and str(qty_raw).strip():
                    qty = parse_qty_value(qty_raw)
            
            # Get unit price
            unit_price = 0
            if unit_col:
                up = clean_numeric(row.get(unit_col, 0))
                if up:
                    unit_price = up
            
            # Get total price
            total_price = 0
            if total_col:
                tp = clean_numeric(row.get(total_col, 0))
                if tp:
                    total_price = tp
            
            # Calculate total if missing
            if total_price == 0 and unit_price > 0:
                total_price = unit_price * qty
            
            # Calculate unit price if missing
            if unit_price == 0 and total_price > 0 and qty > 0:
                unit_price = total_price / qty
            
            items.append({
                'Item_No': no_val,
                'Description': desc_val,
                'Qty': qty,
                'Unit_Price': unit_price,
                'Total_Price': total_price,
                'Is_NIC': is_nic,
                'Source_File': source_file
            })
        except Exception as e:
            continue
    
    return items

def match_items(drawing_no, quotes):
    """Match drawing item to quote by item number."""
    drawing_no_clean = str(drawing_no).strip().lower()
    
    # Exact match
    for q in quotes:
        if str(q.get('Item_No', '')).strip().lower() == drawing_no_clean:
            return q
    
    # Numeric match
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
        
        # Determine status
        if use_categories and cat in [1, 2, 3]:
            status = "Owner Supply"
            issue = supplier_codes.get(cat, "Owner handles")
        elif use_categories and cat == 8:
            status = "Existing"
            issue = "Existing or relocated"
        elif item.get('Description', '').upper() in ('SPARE', '-', 'N/A'):
            status = "N/A"
            issue = "Spare or placeholder"
        elif match:
            if match.get('Is_NIC'):
                status = "NIC"
                issue = "Not In Contract"
            elif match['Qty'] == item['Qty']:
                status = "Quoted"
                issue = None
            else:
                status = "Qty Mismatch"
                issue = f"Drawing: {item['Qty']}, Quote: {match['Qty']}"
        else:
            if use_categories and cat == 7:
                status = "Needs Pricing"
                issue = "Owner supplies - needs install pricing"
            elif use_categories and cat in [5, 6]:
                status = "MISSING"
                issue = "Critical - requires quote"
            else:
                status = "MISSING"
                issue = "Not found in quotes"
        
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
        draw_file = st.file_uploader(
            "Upload drawing schedule (PDF, Excel, CSV)", 
            type=['pdf', 'csv', 'xlsx', 'xls'], 
            key="draw_upload"
        )
        
        if draw_file:
            if draw_file.name != st.session_state.drawing_filename:
                with st.spinner("Processing drawing..."):
                    dfs = parse_uploaded_file(draw_file)
                    if dfs and len(dfs) > 0:
                        combined = max(dfs, key=len)
                        combined = combined.reset_index(drop=True)
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
            no_idx = col_options.index(st.session_state.column_mapping.get('no', '-- Not Used --')) if st.session_state.column_mapping.get('no') in col_options else 0
            no_col = st.selectbox("Item No. Column *", col_options, index=no_idx, key="map_no")
            
            desc_idx = col_options.index(st.session_state.column_mapping.get('description', '-- Not Used --')) if st.session_state.column_mapping.get('description') in col_options else 0
            desc_col = st.selectbox("Description Column *", col_options, index=desc_idx, key="map_desc")
        
        with c2:
            qty_idx = col_options.index(st.session_state.column_mapping.get('qty', '-- Not Used --')) if st.session_state.column_mapping.get('qty') in col_options else 0
            qty_col = st.selectbox("Quantity Column", col_options, index=qty_idx, key="map_qty")
            
            equip_idx = col_options.index(st.session_state.column_mapping.get('equip_num', '-- Not Used --')) if st.session_state.column_mapping.get('equip_num') in col_options else 0
            equip_col = st.selectbox("Equipment/Model # Column", col_options, index=equip_idx, key="map_equip")
        
        with c3:
            st.session_state.use_categories = st.checkbox("Use Category Codes", value=st.session_state.use_categories)
            
            if st.session_state.use_categories:
                cat_idx = col_options.index(st.session_state.column_mapping.get('category', '-- Not Used --')) if st.session_state.column_mapping.get('category') in col_options else 0
                cat_col = st.selectbox("Category Column", col_options, index=cat_idx, key="map_cat")
            else:
                cat_col = '-- Not Used --'
        
        if st.button("✅ Apply Drawing Column Mapping", type="primary"):
            mapping = {}
            if no_col != '-- Not Used --':
                mapping['no'] = no_col
            if desc_col != '-- Not Used --':
                mapping['description'] = desc_col
            if qty_col != '-- Not Used --':
                mapping['qty'] = qty_col
            if equip_col != '-- Not Used --':
                mapping['equip_num'] = equip_col
            if cat_col != '-- Not Used --':
                mapping['category'] = cat_col
            
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
        quote_files = st.file_uploader(
            "Upload quote files (PDF, Excel, CSV)", 
            type=['pdf', 'csv', 'xlsx', 'xls'], 
            accept_multiple_files=True,
            key="quote_upload"
        )
        
        if quote_files:
            for qf in quote_files:
                if qf.name not in st.session_state.quote_dfs:
                    with st.spinner(f"Processing {qf.name}..."):
                        dfs = parse_uploaded_file(qf)
                        if dfs and len(dfs) > 0:
                            combined_df = pd.concat(dfs, ignore_index=True) if len(dfs) > 1 else dfs[0]
                            combined_df = combined_df.reset_index(drop=True)
                            st.session_state.quote_dfs[qf.name] = combined_df
                            st.session_state.quote_mappings[qf.name] = auto_detect_columns(combined_df, 'quote')
                            st.rerun()
                        else:
                            st.error(f"❌ Could not extract data from {qf.name}")
                            st.info("💡 Try using 'Paste Data' option instead - copy data from PDF and paste in CSV format")
    
    else:  # Paste Data
        st.markdown("**Paste quote data in CSV format:**")
        st.caption("Copy from Excel/PDF and paste here. Format: Item, Qty, Description, Sell, Sell Total")
        
        sample_data = """Item,Qty,Description,Sell,Sell Total
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
            st.code(sample_data)
        
        pasted_data = st.text_area("Paste your quote data here:", height=200, placeholder="Item,Qty,Description,Sell,Sell Total\n1,,NIC,,\n2,1 ea,WALK IN,97980.27,97980.27")
        quote_name = st.text_input("Name for this quote:", value="Pasted_Quote")
        
        if st.button("📥 Load Pasted Data", type="primary"):
            if pasted_data.strip():
                try:
                    paste_df = pd.read_csv(io.StringIO(pasted_data))
                    paste_df = clean_dataframe_columns(paste_df)
                    st.session_state.quote_dfs[quote_name] = paste_df
                    st.session_state.quote_mappings[quote_name] = auto_detect_columns(paste_df, 'quote')
                    st.success(f"✅ Loaded {len(paste_df)} rows from pasted data")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error parsing data: {e}")
                    st.info("Make sure data is in CSV format with comma separators")
            else:
                st.warning("Please paste some data first")
    
    # Configure each quote file
    if st.session_state.quote_dfs:
        st.markdown("---")
        st.markdown("**Configure Quote Column Mappings:**")
        
        for filename, qdf in st.session_state.quote_dfs.items():
            with st.expander(f"📄 {filename} ({len(qdf)} rows)", expanded=(filename not in st.session_state.quotes_data)):
                st.dataframe(qdf.head(15), height=180, use_container_width=True)
                
                q_col_options = ['-- Not Used --'] + list(qdf.columns)
                current_map = st.session_state.quote_mappings.get(filename, {})
                
                qc1, qc2 = st.columns(2)
                
                with qc1:
                    q_no_idx = q_col_options.index(current_map.get('no')) if current_map.get('no') in q_col_options else 0
                    q_no_col = st.selectbox("Item No.", q_col_options, index=q_no_idx, key=f"q_no_{filename}")
                    
                    q_desc_idx = q_col_options.index(current_map.get('description')) if current_map.get('description') in q_col_options else 0
                    q_desc_col = st.selectbox("Description", q_col_options, index=q_desc_idx, key=f"q_desc_{filename}")
                    
                    q_qty_idx = q_col_options.index(current_map.get('qty')) if current_map.get('qty') in q_col_options else 0
                    q_qty_col = st.selectbox("Quantity", q_col_options, index=q_qty_idx, key=f"q_qty_{filename}")
                
                with qc2:
                    q_unit_idx = q_col_options.index(current_map.get('unit_price')) if current_map.get('unit_price') in q_col_options else 0
                    q_unit_col = st.selectbox("Unit Price (Sell)", q_col_options, index=q_unit_idx, key=f"q_unit_{filename}")
                    
                    q_total_idx = q_col_options.index(current_map.get('total_price')) if current_map.get('total_price') in q_col_options else 0
                    q_total_col = st.selectbox("Total Price (Sell Total)", q_col_options, index=q_total_idx, key=f"q_total_{filename}")
                
                bcol1, bcol2 = st.columns(2)
                
                with bcol1:
                    if st.button(f"✅ Apply Mapping", key=f"apply_{filename}", type="primary"):
                        q_mapping = {}
                        if q_no_col != '-- Not Used --':
                            q_mapping['no'] = q_no_col
                        if q_desc_col != '-- Not Used --':
                            q_mapping['description'] = q_desc_col
                        if q_qty_col != '-- Not Used --':
                            q_mapping['qty'] = q_qty_col
                        if q_unit_col != '-- Not Used --':
                            q_mapping['unit_price'] = q_unit_col
                        if q_total_col != '-- Not Used --':
                            q_mapping['total_price'] = q_total_col
                        
                        st.session_state.quote_mappings[filename] = q_mapping
                        
                        items = extract_quote_data(qdf, q_mapping, filename)
                        if items:
                            st.session_state.quotes_data[filename] = items
                            nic_count = sum(1 for i in items if i.get('Is_NIC'))
                            total_val = sum(i['Total_Price'] for i in items if not i.get('Is_NIC'))
                            st.success(f"✅ Extracted {len(items)} items ({nic_count} NIC) | ${total_val:,.2f}")
                            st.rerun()
                        else:
                            st.error(f"No items extracted. Check column mapping.")
                
                with bcol2:
                    if st.button(f"🗑️ Remove", key=f"remove_{filename}"):
                        del st.session_state.quote_dfs[filename]
                        if filename in st.session_state.quotes_data:
                            del st.session_state.quotes_data[filename]
                        if filename in st.session_state.quote_mappings:
                            del st.session_state.quote_mappings[filename]
                        st.rerun()
                
                if filename in st.session_state.quotes_data:
                    items = st.session_state.quotes_data[filename]
                    nic_count = sum(1 for i in items if i.get('Is_NIC'))
                    total = sum(i['Total_Price'] for i in items if not i.get('Is_NIC'))
                    st.success(f"✅ {len(items)} items | {nic_count} NIC | Total: ${total:,.2f}")
    
    # Summary
    if st.session_state.quotes_data:
        st.markdown("---")
        st.subheader("📊 Loaded Quotes Summary")
        for fn, qs in st.session_state.quotes_data.items():
            nic_count = sum(1 for q in qs if q.get('Is_NIC'))
            total = sum(q['Total_Price'] for q in qs if not q.get('Is_NIC'))
            st.caption(f"• **{fn}**: {len(qs)} items ({nic_count} NIC) = ${total:,.2f}")
    
    st.markdown("---")
    if st.button("🔄 Reset Everything"):
        st.session_state.drawing_data = None
        st.session_state.drawing_df = None
        st.session_state.quotes_data = {}
        st.session_state.quote_dfs = {}
        st.session_state.quote_mappings = {}
        st.session_state.drawing_filename = None
        st.session_state.column_mapping = {}
        st.rerun()

# ===== TAB 2: Dashboard =====
with tabs[1]:
    if not st.session_state.drawing_data:
        st.warning("⚠️ Please upload and configure drawing first (Tab 1)")
    elif not st.session_state.quotes_data:
        st.warning("⚠️ Please upload and configure quotations (Tab 1)")
        if st.session_state.quote_dfs:
            st.info("💡 You have quote files loaded but haven't applied column mapping. Go to Tab 1 and click 'Apply Mapping'.")
    else:
        all_quotes = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_data(st.session_state.drawing_data, all_quotes, st.session_state.use_categories, st.session_state.supplier_codes)
        
        st.subheader("📊 Coverage Summary")
        
        exclude_statuses = ['Owner Supply', 'Existing', 'N/A']
        actionable = df[~df['Status'].isin(exclude_statuses)]
        
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("✅ Quoted", len(df[df['Status'] == 'Quoted']))
        c2.metric("❌ Missing", len(df[df['Status'] == 'MISSING']))
        c3.metric("⚠️ Qty Mismatch", len(df[df['Status'] == 'Qty Mismatch']))
        c4.metric("🚫 NIC", len(df[df['Status'] == 'NIC']))
        c5.metric("📋 Needs Pricing", len(df[df['Status'] == 'Needs Pricing']))
        
        col1, col2 = st.columns(2)
        col1.metric("💰 Total Quoted Value", f"${df['Total_Price'].sum():,.2f}")
        col2.metric("📦 Total Items", f"{len(df)} ({len(actionable)} actionable)")
        
        ch1, ch2 = st.columns(2)
        
        with ch1:
            vc = df['Status'].value_counts().reset_index()
            vc.columns = ['Status', 'Count']
            colors = {
                'Quoted': '#28a745', 'MISSING': '#dc3545', 'Qty Mismatch': '#ffc107',
                'NIC': '#6f42c1', 'Needs Pricing': '#fd7e14', 'Owner Supply': '#6c757d', 
                'Existing': '#adb5bd', 'N/A': '#e9ecef'
            }
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
        
        col1, col2 = st.columns(2)
        status_opts = df['Status'].unique().tolist()
        filt_status = col1.multiselect("Filter by Status", status_opts, default=status_opts)
        
        fdf = df[df['Status'].isin(filt_status)]
        
        def highlight_row(row):
            colors = {
                'Quoted': 'background-color:#d4edda',
                'MISSING': 'background-color:#f8d7da',
                'Qty Mismatch': 'background-color:#fff3cd',
                'NIC': 'background-color:#e2d5f0',
                'Needs Pricing': 'background-color:#ffe5d0'
            }
            return [colors.get(row['Status'], '')] * len(row)
        
        display_cols = ['Drawing_No', 'Equip_Num', 'Description', 'Drawing_Qty']
        if st.session_state.use_categories:
            display_cols.append('Category')
        display_cols.extend(['Quote_Item_No', 'Quote_Qty', 'Unit_Price', 'Total_Price', 'Status', 'Issue'])
        
        st.dataframe(fdf[display_cols].style.apply(highlight_row, axis=1), height=450, use_container_width=True)
        
        # Critical missing
        st.subheader("🚨 Critical Missing Items")
        if st.session_state.use_categories:
            critical = df[(df['Status'] == 'MISSING') & (df['Category'].isin([5, 6]))]
        else:
            critical = df[df['Status'] == 'MISSING']
        
        if len(critical) > 0:
            st.error(f"⚠️ {len(critical)} items require quotes!")
            st.dataframe(critical[['Drawing_No', 'Equip_Num', 'Description', 'Drawing_Qty']], use_container_width=True)
        else:
            st.success("✅ No critical missing items!")
        
        # NIC Items
        nic_items = df[df['Status'] == 'NIC']
        if len(nic_items) > 0:
            st.subheader("🚫 NIC Items (Not In Contract)")
            st.info(f"{len(nic_items)} items marked as NIC - excluded from vendor scope")
            st.dataframe(nic_items[['Drawing_No', 'Description', 'Drawing_Qty']], use_container_width=True)
    else:
        st.warning("⚠️ Please upload and configure both drawing and quotations first")

# ===== TAB 4: Summary =====
with tabs[3]:
    if st.session_state.drawing_data and st.session_state.quotes_data:
        all_quotes = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_data(st.session_state.drawing_data, all_quotes, st.session_state.use_categories, st.session_state.supplier_codes)
        
        st.subheader("📋 Summary by Status")
        status_summary = df.groupby('Status').agg(
            Items=('Status', 'count'),
            Total_Value=('Total_Price', 'sum')
        ).reset_index()
        status_summary['Total_Value'] = status_summary['Total_Value'].apply(lambda x: f"${x:,.2f}")
        st.dataframe(status_summary, use_container_width=True)
        
        if st.session_state.use_categories and df['Category'].notna().any():
            st.subheader("📋 Summary by Category")
            summary = []
            for code, desc in st.session_state.supplier_codes.items():
                ci = df[df['Category'] == code]
                if len(ci) > 0:
                    summary.append({
                        'Category': code,
                        'Description': desc,
                        'Items': len(ci),
                        'Quoted': len(ci[ci['Status'] == 'Quoted']),
                        'Missing': len(ci[ci['Status'] == 'MISSING']),
                        'NIC': len(ci[ci['Status'] == 'NIC']),
                        'Value': f"${ci['Total_Price'].sum():,.2f}"
                    })
            if summary:
                st.dataframe(pd.DataFrame(summary), use_container_width=True)
        
        st.subheader("📄 Quote Files Summary")
        quote_summary = []
        for fn, qs in st.session_state.quotes_data.items():
            nic_count = sum(1 for q in qs if q.get('Is_NIC'))
            quote_summary.append({
                'File': fn,
                'Total Items': len(qs),
                'NIC Items': nic_count,
                'Priced Items': len(qs) - nic_count,
                'Total Value': f"${sum(q['Total_Price'] for q in qs if not q.get('Is_NIC')):,.2f}"
            })
        st.dataframe(pd.DataFrame(quote_summary), use_container_width=True)
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
            df[df['Status'] == 'Qty Mismatch'].to_excel(writer, sheet_name='Qty Mismatch', index=False)
            pd.DataFrame(all_quotes).to_excel(writer, sheet_name='All Quotes Raw', index=False)
        out.seek(0)
        
        st.download_button(
            "📥 Download Full Analysis Report",
            out,
            f"Analysis_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
        st.markdown("---")
        col1, col2 = st.columns(2)
        
        with col1:
            if st.session_state.drawing_data:
                st.markdown("**Drawing Data**")
                eq_df = pd.DataFrame(st.session_state.drawing_data)
                out_eq = io.BytesIO()
                eq_df.to_excel(out_eq, index=False)
                out_eq.seek(0)
                st.download_button("📥 Download Drawing Data", out_eq, f"Drawing_{datetime.now().strftime('%Y%m%d')}.xlsx")
        
        with col2:
            if st.session_state.quotes_data:
                st.markdown("**Quote Data**")
                all_q = [q for qs in st.session_state.quotes_data.values() for q in qs]
                q_df = pd.DataFrame(all_q)
                out_q = io.BytesIO()
                q_df.to_excel(out_q, index=False)
                out_q.seek(0)
                st.download_button("📥 Download Quote Data", out_q, f"Quotes_{datetime.now().strftime('%Y%m%d')}.xlsx")
    else:
        st.warning("⚠️ Please upload and configure both drawing and quotations first")

st.markdown("---")
st.caption("Universal Drawing Quote Analyzer v10.0 | NIC = Not In Contract")
