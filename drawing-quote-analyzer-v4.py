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
    1: "Owner Supply / Owner Install", 2: "Owner Supply / Owner Install (Special)",
    3: "Owner Supply / Owner Install (Other)", 4: "Owner Supply / Vendor Install",
    5: "Contractor Supply / Contractor Install", 6: "Contractor Supply / Vendor Install",
    7: "Owner Supply / Contractor Install", 8: "Existing / Relocated"
}

for key, default in [('drawing_data', None), ('drawing_df', None), ('quotes_data', {}),
    ('quote_dfs', {}), ('quote_mappings', {}), ('drawing_filename', None),
    ('column_mapping', {}), ('supplier_codes', DEFAULT_SUPPLIER_CODES.copy()), ('use_categories', True)]:
    if key not in st.session_state:
        st.session_state[key] = default

def clean_df_columns(df):
    df = df.copy()
    new_cols = []
    seen = {}
    for i, c in enumerate(df.columns):
        col = str(c).strip() if pd.notna(c) and str(c).strip() else f'Col_{i}'
        if col in seen:
            seen[col] += 1
            col = f"{col}_{seen[col]}"
        else:
            seen[col] = 0
        new_cols.append(col)
    df.columns = new_cols
    return df.dropna(how='all')

def parse_qty(val):
    if pd.isna(val): return 1
    val_str = str(val).strip().lower()
    match = re.search(r'(\d+)\s*ea', val_str)
    if match: return int(match.group(1))
    match = re.search(r'^(\d+)', val_str)
    if match: return int(match.group(1))
    return 1

def clean_price(val):
    if pd.isna(val): return 0.0
    val_str = re.sub(r'[,$\s]', '', str(val).strip())
    try: return float(val_str) if val_str and val_str.replace('.','').replace('-','').isdigit() else 0.0
    except: return 0.0

def extract_pdf_method1(uploaded_file):
    """Method 1: Standard table extraction"""
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
                        if len(df) > 0:
                            all_tables.append(df)
        if all_tables:
            return pd.concat(all_tables, ignore_index=True)
    except Exception as e:
        st.warning(f"Method 1 failed: {e}")
    return None

def extract_pdf_method2(uploaded_file):
    """Method 2: Extract tables with different settings"""
    uploaded_file.seek(0)
    all_rows = []
    try:
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables(table_settings={
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                    "snap_tolerance": 5,
                })
                for table in tables:
                    if table:
                        for row in table:
                            if row and any(c and str(c).strip() for c in row if c):
                                all_rows.append([str(c).strip() if c else '' for c in row])
        if all_rows and len(all_rows) > 1:
            # Find most common length
            from collections import Counter
            lengths = [len(r) for r in all_rows]
            common_len = Counter(lengths).most_common(1)[0][0]
            normalized = [r for r in all_rows if len(r) == common_len]
            if normalized:
                df = pd.DataFrame(normalized[1:], columns=normalized[0])
                return df
    except Exception as e:
        st.warning(f"Method 2 failed: {e}")
    return None

def extract_pdf_method3(uploaded_file):
    """Method 3: Text-based extraction for quote format"""
    uploaded_file.seek(0)
    items = []
    try:
        with pdfplumber.open(uploaded_file) as pdf:
            full_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    full_text += text + "\n"
            
            # Parse line by line looking for item patterns
            lines = full_text.split('\n')
            current_item = None
            
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                
                # Pattern: starts with number, has qty like "1 ea" or "2 ea", has price
                # Example: "2 1 ea WALK IN $97,980.27 $97,980.27"
                # Or: "1 NIC"
                # Or: "10 2 ea STAINLESS $2,206.08 $4,412.16"
                
                # Check for NIC pattern: "1 NIC" or "11-23 NIC"
                nic_match = re.match(r'^(\d+(?:-\d+)?)\s+NIC\s*$', line, re.IGNORECASE)
                if nic_match:
                    items.append({
                        'Item': nic_match.group(1),
                        'Qty': '',
                        'Description': 'NIC',
                        'Sell': '',
                        'Sell Total': ''
                    })
                    continue
                
                # Check for item with pricing pattern
                # Pattern: Item# Qty Description Price Price
                item_match = re.match(r'^(\d+)\s+(\d+\s*ea)\s+(.+?)\s+\$?([\d,]+\.?\d*)\s+\$?([\d,]+\.?\d*)\s*$', line, re.IGNORECASE)
                if item_match:
                    items.append({
                        'Item': item_match.group(1),
                        'Qty': item_match.group(2),
                        'Description': item_match.group(3).strip(),
                        'Sell': item_match.group(4),
                        'Sell Total': item_match.group(5)
                    })
                    continue
                
                # Check for item without pricing (description continues)
                item_match2 = re.match(r'^(\d+)\s+(\d+\s*ea)\s+(.+)$', line, re.IGNORECASE)
                if item_match2 and not re.search(r'\$', line):
                    items.append({
                        'Item': item_match2.group(1),
                        'Qty': item_match2.group(2),
                        'Description': item_match2.group(3).strip(),
                        'Sell': '',
                        'Sell Total': ''
                    })
            
            if items:
                return pd.DataFrame(items)
    except Exception as e:
        st.warning(f"Method 3 failed: {e}")
    return None

def extract_pdf_all_text(uploaded_file):
    """Extract all text for debugging"""
    uploaded_file.seek(0)
    try:
        with pdfplumber.open(uploaded_file) as pdf:
            text = ""
            for i, page in enumerate(pdf.pages[:3]):  # First 3 pages
                t = page.extract_text()
                if t:
                    text += f"=== PAGE {i+1} ===\n{t}\n\n"
            return text
    except Exception as e:
        return f"Error: {e}"

def parse_excel(uploaded_file):
    try:
        uploaded_file.seek(0)
        xl = pd.ExcelFile(uploaded_file)
        dfs = [clean_df_columns(pd.read_excel(xl, sheet_name=s)) for s in xl.sheet_names]
        return [d for d in dfs if len(d) > 0] or None
    except Exception as e:
        st.warning(f"Excel error: {e}")
        return None

def parse_csv(uploaded_file):
    try:
        uploaded_file.seek(0)
        df = clean_df_columns(pd.read_csv(uploaded_file))
        return [df] if len(df) > 0 else None
    except Exception as e:
        st.warning(f"CSV error: {e}")
        return None

def parse_file(uploaded_file):
    ext = uploaded_file.name.split('.')[-1].lower()
    if ext == 'pdf':
        # Try multiple methods
        df = extract_pdf_method1(uploaded_file)
        if df is not None and len(df) > 0:
            return [clean_df_columns(df)]
        
        df = extract_pdf_method2(uploaded_file)
        if df is not None and len(df) > 0:
            return [clean_df_columns(df)]
        
        df = extract_pdf_method3(uploaded_file)
        if df is not None and len(df) > 0:
            return [clean_df_columns(df)]
        
        return None
    elif ext in ['xlsx', 'xls']:
        return parse_excel(uploaded_file)
    elif ext == 'csv':
        return parse_csv(uploaded_file)
    return None

def auto_detect_cols(df, file_type='drawing'):
    cols_lower = {c: c.lower().strip() for c in df.columns}
    patterns = {
        'drawing': {'no': ['no', 'item', '#'], 'description': ['description', 'desc', 'equipment'], 
                   'qty': ['qty', 'quantity'], 'category': ['category', 'cat', 'code'], 'equip_num': ['equip', 'model']},
        'quote': {'no': ['item', 'no', '#'], 'description': ['description', 'desc'], 
                 'qty': ['qty', 'quantity'], 'unit_price': ['sell', 'price', 'unit'], 'total_price': ['total', 'sell total', 'ext']}
    }[file_type]
    
    found = {}
    for key, opts in patterns.items():
        for col, col_low in cols_lower.items():
            if any(o in col_low for o in opts):
                found[key] = col
                break
    return found

def extract_drawing_data(df, col_map):
    items = []
    no_col, desc_col = col_map.get('no'), col_map.get('description')
    if not no_col or not desc_col: return None
    
    for _, row in df.iterrows():
        no_val = str(row.get(no_col, '')).strip()
        desc_val = str(row.get(desc_col, '')).strip()
        if not no_val or no_val.lower() in ('nan', '', 'no', 'item'): continue
        if not desc_val or desc_val.lower() in ('nan', '', 'description'): continue
        
        qty = parse_qty(row.get(col_map.get('qty'), 1)) if col_map.get('qty') else 1
        cat = int(clean_price(row.get(col_map.get('category'), ''))) if col_map.get('category') and clean_price(row.get(col_map.get('category'), '')) else None
        equip = str(row.get(col_map.get('equip_num'), '-')).strip() if col_map.get('equip_num') else '-'
        if equip.lower() in ('nan', '', 'none'): equip = '-'
        
        items.append({'No': no_val, 'Equip_Num': equip, 'Description': desc_val, 'Qty': qty, 'Category': cat})
    return items if items else None

def extract_quote_data(df, col_map, source_file):
    items = []
    no_col, desc_col = col_map.get('no'), col_map.get('description')
    
    for _, row in df.iterrows():
        no_val = str(row.get(no_col, '')).strip() if no_col else ''
        desc_val = str(row.get(desc_col, '')).strip() if desc_col else ''
        if no_val.lower() in ('nan', 'none'): no_val = ''
        if desc_val.lower() in ('nan', 'none'): desc_val = ''
        if not no_val and not desc_val: continue
        if no_val.lower() in ('item', 'no') or desc_val.lower() == 'description': continue
        
        is_nic = 'NIC' in desc_val.upper() or desc_val.upper().strip() == 'NIC'
        qty = parse_qty(row.get(col_map.get('qty'), '')) if col_map.get('qty') else 1
        unit_price = clean_price(row.get(col_map.get('unit_price'), 0)) if col_map.get('unit_price') else 0
        total_price = clean_price(row.get(col_map.get('total_price'), 0)) if col_map.get('total_price') else 0
        
        if total_price == 0 and unit_price > 0: total_price = unit_price * qty
        if unit_price == 0 and total_price > 0 and qty > 0: unit_price = total_price / qty
        
        items.append({'Item_No': no_val, 'Description': desc_val, 'Qty': qty, 'Unit_Price': unit_price, 
                     'Total_Price': total_price, 'Is_NIC': is_nic, 'Source_File': source_file})
    return items

def match_item(drawing_no, quotes):
    no_clean = str(drawing_no).strip().lower()
    for q in quotes:
        if str(q.get('Item_No', '')).strip().lower() == no_clean: return q
    try:
        draw_num = int(re.sub(r'[^0-9]', '', no_clean))
        for q in quotes:
            try:
                if draw_num == int(re.sub(r'[^0-9]', '', str(q.get('Item_No', '')))): return q
            except: pass
    except: pass
    return None

def analyze_data(drawing_items, quotes, use_categories=True, supplier_codes=None):
    if supplier_codes is None: supplier_codes = DEFAULT_SUPPLIER_CODES
    results = []
    for item in drawing_items:
        match = match_item(item['No'], quotes)
        cat = item.get('Category')
        
        if use_categories and cat in [1, 2, 3]: status, issue = "Owner Supply", supplier_codes.get(cat, "Owner")
        elif use_categories and cat == 8: status, issue = "Existing", "Existing/relocated"
        elif item.get('Description', '').upper() in ('SPARE', '-', 'N/A'): status, issue = "N/A", "Spare"
        elif match:
            if match.get('Is_NIC'): status, issue = "NIC", "Not In Contract"
            elif match['Qty'] == item['Qty']: status, issue = "Quoted", None
            else: status, issue = "Qty Mismatch", f"Draw:{item['Qty']} vs Quote:{match['Qty']}"
        else:
            if use_categories and cat == 7: status, issue = "Needs Pricing", "Owner supply - needs install"
            elif use_categories and cat in [5, 6]: status, issue = "MISSING", "Critical - needs quote"
            else: status, issue = "MISSING", "Not in quotes"
        
        results.append({
            'Drawing_No': item['No'], 'Equip_Num': item.get('Equip_Num', '-'), 'Description': item['Description'],
            'Drawing_Qty': item['Qty'], 'Category': cat, 'Category_Desc': supplier_codes.get(cat, '-') if cat and use_categories else '-',
            'Quote_Item_No': match['Item_No'] if match else '-', 'Quote_Qty': match['Qty'] if match else 0,
            'Unit_Price': match['Unit_Price'] if match else 0, 'Total_Price': match['Total_Price'] if match and not match.get('Is_NIC') else 0,
            'Quote_Source': match['Source_File'] if match else '-', 'Status': status, 'Issue': issue
        })
    return pd.DataFrame(results)

# ===== UI =====
st.markdown("## 📊 Drawing vs Quote Analyzer")
if not PDF_SUPPORT: st.warning("⚠️ Install pdfplumber: `pip install pdfplumber`")

tabs = st.tabs(["📁 Upload", "📊 Dashboard", "🔍 Analysis", "📋 Summary", "💾 Export"])

with tabs[0]:
    # Drawing Section
    st.subheader("1️⃣ Drawing/Schedule")
    draw_file = st.file_uploader("Upload drawing", type=['pdf', 'csv', 'xlsx', 'xls'], key="draw")
    
    if draw_file and draw_file.name != st.session_state.drawing_filename:
        dfs = parse_file(draw_file)
        if dfs:
            st.session_state.drawing_df = max(dfs, key=len).reset_index(drop=True)
            st.session_state.drawing_filename = draw_file.name
            st.session_state.column_mapping = auto_detect_cols(st.session_state.drawing_df, 'drawing')
            st.session_state.drawing_data = None
            st.rerun()
        else: st.error("Could not extract drawing data")
    
    if st.session_state.drawing_filename: st.success(f"✅ {st.session_state.drawing_filename}")
    
    if st.session_state.drawing_df is not None:
        df = st.session_state.drawing_df
        opts = ['-- Not Used --'] + list(df.columns)
        mapping = st.session_state.column_mapping
        
        with st.expander("Preview & Map Columns", expanded=True):
            st.dataframe(df.head(10), height=150, use_container_width=True)
            c1, c2, c3 = st.columns(3)
            with c1:
                no_col = st.selectbox("Item No*", opts, index=opts.index(mapping.get('no')) if mapping.get('no') in opts else 0)
                desc_col = st.selectbox("Description*", opts, index=opts.index(mapping.get('description')) if mapping.get('description') in opts else 0)
            with c2:
                qty_col = st.selectbox("Qty", opts, index=opts.index(mapping.get('qty')) if mapping.get('qty') in opts else 0)
                equip_col = st.selectbox("Equip #", opts, index=opts.index(mapping.get('equip_num')) if mapping.get('equip_num') in opts else 0)
            with c3:
                st.session_state.use_categories = st.checkbox("Use Categories", st.session_state.use_categories)
                cat_col = st.selectbox("Category", opts, index=opts.index(mapping.get('category')) if mapping.get('category') in opts else 0) if st.session_state.use_categories else '-- Not Used --'
            
            if st.button("✅ Apply Drawing Mapping", type="primary"):
                new_map = {k: v for k, v in [('no', no_col), ('description', desc_col), ('qty', qty_col), ('equip_num', equip_col), ('category', cat_col)] if v != '-- Not Used --'}
                st.session_state.column_mapping = new_map
                items = extract_drawing_data(df, new_map)
                if items:
                    st.session_state.drawing_data = items
                    st.success(f"✅ {len(items)} items extracted")
                    st.rerun()
                else: st.error("No items extracted")
    
    if st.session_state.drawing_data:
        st.caption(f"📋 {len(st.session_state.drawing_data)} drawing items loaded")
    
    # Quote Section
    st.markdown("---")
    st.subheader("2️⃣ Quotations")
    
    quote_method = st.radio("Input Method:", ["Upload File", "Paste Data (CSV format)"], horizontal=True)
    
    if quote_method == "Upload File":
        quote_files = st.file_uploader("Upload quotes", type=['pdf', 'csv', 'xlsx', 'xls'], accept_multiple_files=True, key="quotes")
        
        if quote_files:
            for qf in quote_files:
                if qf.name not in st.session_state.quote_dfs:
                    with st.spinner(f"Processing {qf.name}..."):
                        dfs = parse_file(qf)
                        if dfs:
                            combined = pd.concat(dfs, ignore_index=True) if len(dfs) > 1 else dfs[0]
                            st.session_state.quote_dfs[qf.name] = clean_df_columns(combined)
                            st.session_state.quote_mappings[qf.name] = auto_detect_cols(combined, 'quote')
                            st.rerun()
                        else:
                            st.error(f"❌ Could not extract from {qf.name}")
                            # Show debug info
                            if qf.name.endswith('.pdf') and PDF_SUPPORT:
                                with st.expander("🔍 Debug: Raw PDF Text (first 3 pages)"):
                                    text = extract_pdf_all_text(qf)
                                    st.text_area("Raw text:", text, height=300)
                                    st.info("💡 If you see the data above, try the 'Paste Data' option instead")
    
    else:  # Paste Data
        st.markdown("**Paste quote data in CSV format:**")
        st.caption("Format: Item,Qty,Description,Sell,Sell Total (one row per item)")
        
        sample = """Item,Qty,Description,Sell,Sell Total
1,,NIC,,
2,1 ea,WALK IN,97980.27,97980.27
3,1 ea,WALK IN,70727.09,70727.09
10,2 ea,STAINLESS,2206.08,4412.16
11-23,,NIC,,
24,1 ea,INGREDIENT BIN,386.48,386.48"""
        
        pasted_data = st.text_area("Paste data here:", value="", height=200, placeholder=sample)
        paste_name = st.text_input("Quote name:", value="Pasted_Quote")
        
        if st.button("📥 Load Pasted Data") and pasted_data.strip():
            try:
                df = pd.read_csv(io.StringIO(pasted_data))
                df = clean_df_columns(df)
                st.session_state.quote_dfs[paste_name] = df
                st.session_state.quote_mappings[paste_name] = auto_detect_cols(df, 'quote')
                st.success(f"✅ Loaded {len(df)} rows")
                st.rerun()
            except Exception as e:
                st.error(f"Error parsing: {e}")
    
    # Configure Quotes
    if st.session_state.quote_dfs:
        st.markdown("**Configure Quote Columns:**")
        for fname, qdf in st.session_state.quote_dfs.items():
            with st.expander(f"📄 {fname} ({len(qdf)} rows)", expanded=(fname not in st.session_state.quotes_data)):
                st.dataframe(qdf.head(15), height=150, use_container_width=True)
                
                opts = ['-- Not Used --'] + list(qdf.columns)
                qmap = st.session_state.quote_mappings.get(fname, {})
                
                c1, c2 = st.columns(2)
                with c1:
                    q_no = st.selectbox("Item No", opts, index=opts.index(qmap.get('no')) if qmap.get('no') in opts else 0, key=f"n_{fname}")
                    q_desc = st.selectbox("Description", opts, index=opts.index(qmap.get('description')) if qmap.get('description') in opts else 0, key=f"d_{fname}")
                    q_qty = st.selectbox("Qty", opts, index=opts.index(qmap.get('qty')) if qmap.get('qty') in opts else 0, key=f"q_{fname}")
                with c2:
                    q_unit = st.selectbox("Unit Price", opts, index=opts.index(qmap.get('unit_price')) if qmap.get('unit_price') in opts else 0, key=f"u_{fname}")
                    q_total = st.selectbox("Total Price", opts, index=opts.index(qmap.get('total_price')) if qmap.get('total_price') in opts else 0, key=f"t_{fname}")
                
                bc1, bc2 = st.columns(2)
                with bc1:
                    if st.button("✅ Apply", key=f"a_{fname}", type="primary"):
                        new_qmap = {k: v for k, v in [('no', q_no), ('description', q_desc), ('qty', q_qty), ('unit_price', q_unit), ('total_price', q_total)] if v != '-- Not Used --'}
                        st.session_state.quote_mappings[fname] = new_qmap
                        items = extract_quote_data(qdf, new_qmap, fname)
                        if items:
                            st.session_state.quotes_data[fname] = items
                            nic = sum(1 for i in items if i.get('Is_NIC'))
                            st.success(f"✅ {len(items)} items ({nic} NIC)")
                            st.rerun()
                        else: st.error("No items extracted")
                with bc2:
                    if st.button("🗑️ Remove", key=f"r_{fname}"):
                        del st.session_state.quote_dfs[fname]
                        st.session_state.quotes_data.pop(fname, None)
                        st.session_state.quote_mappings.pop(fname, None)
                        st.rerun()
                
                if fname in st.session_state.quotes_data:
                    items = st.session_state.quotes_data[fname]
                    nic = sum(1 for i in items if i.get('Is_NIC'))
                    total = sum(i['Total_Price'] for i in items if not i.get('Is_NIC'))
                    st.success(f"✅ Loaded: {len(items)} items | {nic} NIC | ${total:,.2f}")
    
    if st.session_state.quotes_data:
        st.markdown("---")
        for fn, qs in st.session_state.quotes_data.items():
            nic = sum(1 for q in qs if q.get('Is_NIC'))
            total = sum(q['Total_Price'] for q in qs if not q.get('Is_NIC'))
            st.caption(f"✅ {fn}: {len(qs)} items ({nic} NIC) = ${total:,.2f}")
    
    if st.button("🔄 Reset All"):
        for k in ['drawing_data', 'drawing_df', 'quotes_data', 'quote_dfs', 'quote_mappings', 'drawing_filename', 'column_mapping']:
            st.session_state[k] = {} if 'data' in k or 'dfs' in k or 'map' in k else None
        st.rerun()

with tabs[1]:
    if not st.session_state.drawing_data: st.warning("⚠️ Configure drawing first")
    elif not st.session_state.quotes_data: st.warning("⚠️ Configure quotes first")
    else:
        all_quotes = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_data(st.session_state.drawing_data, all_quotes, st.session_state.use_categories, st.session_state.supplier_codes)
        
        st.subheader("📊 Coverage Summary")
        actionable = df[~df['Status'].isin(['Owner Supply', 'Existing', 'N/A'])]
        
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("✅ Quoted", len(df[df['Status'] == 'Quoted']))
        c2.metric("❌ Missing", len(df[df['Status'] == 'MISSING']))
        c3.metric("⚠️ Mismatch", len(df[df['Status'] == 'Qty Mismatch']))
        c4.metric("🚫 NIC", len(df[df['Status'] == 'NIC']))
        c5.metric("📋 Needs Price", len(df[df['Status'] == 'Needs Pricing']))
        
        col1, col2 = st.columns(2)
        col1.metric("💰 Quoted Value", f"${df['Total_Price'].sum():,.2f}")
        col2.metric("📦 Total Items", f"{len(df)} ({len(actionable)} actionable)")
        
        ch1, ch2 = st.columns(2)
        with ch1:
            vc = df['Status'].value_counts().reset_index()
            vc.columns = ['Status', 'Count']
            colors = {'Quoted': '#28a745', 'MISSING': '#dc3545', 'Qty Mismatch': '#ffc107', 'NIC': '#6f42c1', 'Needs Pricing': '#fd7e14', 'Owner Supply': '#6c757d', 'Existing': '#adb5bd', 'N/A': '#e9ecef'}
            fig = px.pie(vc, values='Count', names='Status', color='Status', color_discrete_map=colors, title="Status Distribution")
            st.plotly_chart(fig, use_container_width=True)
        with ch2:
            fig2 = px.bar(vc, x='Status', y='Count', color='Status', color_discrete_map=colors, title="Items by Status")
            st.plotly_chart(fig2, use_container_width=True)

with tabs[2]:
    if st.session_state.drawing_data and st.session_state.quotes_data:
        all_quotes = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_data(st.session_state.drawing_data, all_quotes, st.session_state.use_categories, st.session_state.supplier_codes)
        
        filt = st.multiselect("Filter Status", df['Status'].unique().tolist(), default=df['Status'].unique().tolist())
        fdf = df[df['Status'].isin(filt)]
        
        def hl(row):
            c = {'Quoted': '#d4edda', 'MISSING': '#f8d7da', 'Qty Mismatch': '#fff3cd', 'NIC': '#e2d5f0', 'Needs Pricing': '#ffe5d0'}
            return [f'background-color:{c.get(row["Status"], "")}'] * len(row)
        
        cols = ['Drawing_No', 'Description', 'Drawing_Qty', 'Quote_Item_No', 'Quote_Qty', 'Total_Price', 'Status', 'Issue']
        st.dataframe(fdf[cols].style.apply(hl, axis=1), height=400, use_container_width=True)
        
        critical = df[df['Status'] == 'MISSING']
        if len(critical) > 0:
            st.error(f"🚨 {len(critical)} MISSING items!")
            st.dataframe(critical[['Drawing_No', 'Description', 'Drawing_Qty']], use_container_width=True)
    else: st.warning("⚠️ Configure data first")

with tabs[3]:
    if st.session_state.drawing_data and st.session_state.quotes_data:
        all_quotes = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_data(st.session_state.drawing_data, all_quotes, st.session_state.use_categories, st.session_state.supplier_codes)
        summary = df.groupby('Status').agg(Items=('Status', 'count'), Value=('Total_Price', 'sum')).reset_index()
        summary['Value'] = summary['Value'].apply(lambda x: f"${x:,.2f}")
        st.dataframe(summary, use_container_width=True)
    else: st.warning("⚠️ Configure data first")

with tabs[4]:
    if st.session_state.drawing_data and st.session_state.quotes_data:
        all_quotes = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_data(st.session_state.drawing_data, all_quotes, st.session_state.use_categories, st.session_state.supplier_codes)
        
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as w:
            df.to_excel(w, sheet_name='Full', index=False)
            df[df['Status'] == 'MISSING'].to_excel(w, sheet_name='Missing', index=False)
            df[df['Status'] == 'NIC'].to_excel(w, sheet_name='NIC', index=False)
        out.seek(0)
        st.download_button("📥 Download Report", out, f"Analysis_{datetime.now().strftime('%Y%m%d')}.xlsx", type="primary")
    else: st.warning("⚠️ Configure data first")

st.caption("v9.0 | NIC = Not In Contract")
