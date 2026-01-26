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

# Session state initialization
for key, default in [
    ('drawing_data', None), ('drawing_df', None), ('quotes_data', {}),
    ('quote_dfs', {}), ('quote_mappings', {}), ('drawing_filename', None),
    ('column_mapping', {}), ('supplier_codes', DEFAULT_SUPPLIER_CODES.copy()),
    ('use_categories', True)
]:
    if key not in st.session_state:
        st.session_state[key] = default

def clean_df_columns(df):
    """Clean DataFrame columns."""
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
    """Parse quantity from various formats like '1 ea', '2', etc."""
    if pd.isna(val):
        return 1
    val_str = str(val).strip().lower()
    # Handle "X ea" format
    match = re.search(r'(\d+)\s*ea', val_str)
    if match:
        return int(match.group(1))
    # Handle plain numbers
    match = re.search(r'^(\d+)', val_str)
    if match:
        return int(match.group(1))
    return 1

def clean_price(val):
    """Clean price value."""
    if pd.isna(val):
        return 0.0
    val_str = str(val).strip()
    val_str = re.sub(r'[,$]', '', val_str)
    val_str = re.sub(r'[^\d.\-]', '', val_str)
    try:
        return float(val_str) if val_str else 0.0
    except:
        return 0.0

def extract_pdf_tables(uploaded_file):
    """Extract tables from PDF with multiple strategies."""
    if not PDF_SUPPORT:
        return None
    
    uploaded_file.seek(0)
    all_rows = []
    
    try:
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table:
                        for row in table:
                            if row and any(cell and str(cell).strip() for cell in row if cell):
                                cleaned = [str(c).strip() if c else '' for c in row]
                                all_rows.append(cleaned)
        
        if not all_rows:
            return None
        
        # Find most common row length
        from collections import Counter
        lengths = [len(r) for r in all_rows if len(r) >= 3]
        if not lengths:
            return None
        common_len = Counter(lengths).most_common(1)[0][0]
        
        # Normalize rows
        normalized = []
        for row in all_rows:
            if len(row) >= common_len - 1:
                if len(row) < common_len:
                    row = row + [''] * (common_len - len(row))
                elif len(row) > common_len:
                    row = row[:common_len]
                normalized.append(row)
        
        if not normalized:
            return None
        
        # Check if first row is header
        first = normalized[0]
        header_words = ['item', 'qty', 'description', 'sell', 'total', 'price', 'no', 'quantity']
        is_header = any(any(hw in str(c).lower() for hw in header_words) for c in first if c)
        
        if is_header:
            headers = [str(h).strip() if h else f'Col_{i}' for i, h in enumerate(first)]
            data = normalized[1:]
        else:
            headers = [f'Column_{i}' for i in range(common_len)]
            data = normalized
        
        df = pd.DataFrame(data, columns=headers)
        return [clean_df_columns(df)] if len(df) > 0 else None
        
    except Exception as e:
        st.warning(f"PDF error: {e}")
        return None

def parse_excel(uploaded_file):
    """Parse Excel file."""
    try:
        uploaded_file.seek(0)
        xl = pd.ExcelFile(uploaded_file)
        dfs = []
        for sheet in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name=sheet)
            df = clean_df_columns(df)
            if len(df) > 0:
                dfs.append(df)
        return dfs if dfs else None
    except Exception as e:
        st.warning(f"Excel error: {e}")
        return None

def parse_csv(uploaded_file):
    """Parse CSV file."""
    try:
        uploaded_file.seek(0)
        df = pd.read_csv(uploaded_file)
        df = clean_df_columns(df)
        return [df] if len(df) > 0 else None
    except Exception as e:
        st.warning(f"CSV error: {e}")
        return None

def parse_file(uploaded_file):
    """Parse any supported file."""
    ext = uploaded_file.name.split('.')[-1].lower()
    if ext == 'pdf':
        return extract_pdf_tables(uploaded_file)
    elif ext in ['xlsx', 'xls']:
        return parse_excel(uploaded_file)
    elif ext == 'csv':
        return parse_csv(uploaded_file)
    return None

def auto_detect_cols(df, file_type='drawing'):
    """Auto-detect column mappings."""
    cols_lower = {c: c.lower().strip() for c in df.columns}
    
    patterns = {
        'drawing': {
            'no': ['no', 'no.', 'item', 'item #', 'item no', 'number', '#'],
            'description': ['description', 'desc', 'equipment', 'name', 'material'],
            'qty': ['qty', 'qty.', 'quantity', 'count', 'amount'],
            'category': ['category', 'cat', 'supplier code', 'code', 'type'],
            'equip_num': ['equipment number', 'equip num', 'model', 'part no'],
        },
        'quote': {
            'no': ['item', 'no', 'no.', 'item #', '#', 'line'],
            'description': ['description', 'desc', 'equipment', 'name', 'product'],
            'qty': ['qty', 'qty.', 'quantity', 'count'],
            'unit_price': ['sell', 'unit price', 'price', 'rate', 'unit'],
            'total_price': ['sell total', 'total', 'ext', 'extended', 'amount'],
        }
    }[file_type]
    
    found = {}
    for key, opts in patterns.items():
        for col, col_low in cols_lower.items():
            if col_low in opts or any(o in col_low for o in opts):
                found[key] = col
                break
    return found

def extract_drawing_data(df, col_map):
    """Extract drawing data."""
    items = []
    no_col = col_map.get('no')
    desc_col = col_map.get('description')
    qty_col = col_map.get('qty')
    cat_col = col_map.get('category')
    equip_col = col_map.get('equip_num')
    
    if not no_col or not desc_col:
        return None
    
    for _, row in df.iterrows():
        try:
            no_val = str(row.get(no_col, '')).strip()
            desc_val = str(row.get(desc_col, '')).strip()
            
            if not no_val or no_val.lower() in ('nan', '', 'no', 'no.', 'item', 'none'):
                continue
            if not desc_val or desc_val.lower() in ('nan', '', 'description', 'none'):
                continue
            
            qty = parse_qty(row.get(qty_col, 1)) if qty_col else 1
            cat = None
            if cat_col:
                cat_val = clean_price(row.get(cat_col, ''))
                if cat_val:
                    cat = int(cat_val)
            
            equip_num = '-'
            if equip_col:
                en = str(row.get(equip_col, '')).strip()
                if en and en.lower() not in ('nan', '', '-', 'none'):
                    equip_num = en
            
            items.append({
                'No': no_val, 'Equip_Num': equip_num,
                'Description': desc_val, 'Qty': qty, 'Category': cat
            })
        except:
            continue
    
    return items if items else None

def extract_quote_data(df, col_map, source_file):
    """Extract quote data with NIC handling."""
    items = []
    no_col = col_map.get('no')
    desc_col = col_map.get('description')
    qty_col = col_map.get('qty')
    unit_col = col_map.get('unit_price')
    total_col = col_map.get('total_price')
    
    for _, row in df.iterrows():
        try:
            # Get item number
            no_val = str(row.get(no_col, '')).strip() if no_col else ''
            if no_val.lower() in ('nan', 'none'):
                no_val = ''
            
            # Get description
            desc_val = str(row.get(desc_col, '')).strip() if desc_col else ''
            if desc_val.lower() in ('nan', 'none'):
                desc_val = ''
            
            # Skip empty rows or headers
            if not no_val and not desc_val:
                continue
            if no_val.lower() in ('item', 'no', 'no.', '#') or desc_val.lower() == 'description':
                continue
            
            # Check for NIC (Not In Contract)
            is_nic = 'NIC' in desc_val.upper() or desc_val.upper().strip() == 'NIC'
            
            # Get quantity
            qty = 1
            if qty_col:
                qty_raw = row.get(qty_col, '')
                if pd.notna(qty_raw) and str(qty_raw).strip():
                    qty = parse_qty(qty_raw)
            
            # Get prices
            unit_price = clean_price(row.get(unit_col, 0)) if unit_col else 0
            total_price = clean_price(row.get(total_col, 0)) if total_col else 0
            
            # Calculate missing values
            if total_price == 0 and unit_price > 0:
                total_price = unit_price * qty
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
        except:
            continue
    
    return items

def match_item(drawing_no, quotes):
    """Match drawing item to quote."""
    no_clean = str(drawing_no).strip().lower()
    
    # Exact match
    for q in quotes:
        if str(q.get('Item_No', '')).strip().lower() == no_clean:
            return q
    
    # Numeric match
    try:
        draw_num = int(re.sub(r'[^0-9]', '', no_clean))
        for q in quotes:
            try:
                q_num = int(re.sub(r'[^0-9]', '', str(q.get('Item_No', '')).strip()))
                if draw_num == q_num:
                    return q
            except:
                pass
    except:
        pass
    
    return None

def analyze_data(drawing_items, quotes, use_categories=True, supplier_codes=None):
    """Analyze drawing vs quotes."""
    if supplier_codes is None:
        supplier_codes = DEFAULT_SUPPLIER_CODES
    
    results = []
    for item in drawing_items:
        match = match_item(item['No'], quotes)
        cat = item.get('Category')
        
        # Determine status
        if use_categories and cat in [1, 2, 3]:
            status, issue = "Owner Supply", supplier_codes.get(cat, "Owner handles")
        elif use_categories and cat == 8:
            status, issue = "Existing", "Existing or relocated"
        elif item.get('Description', '').upper() in ('SPARE', '-', 'N/A'):
            status, issue = "N/A", "Spare or placeholder"
        elif match:
            if match.get('Is_NIC'):
                status, issue = "NIC", "Not In Contract - excluded from quote"
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
        
        results.append({
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
    
    return pd.DataFrame(results)

# ===== UI =====
st.markdown("## 📊 Drawing vs Quote Analyzer")
st.caption("Compare equipment schedules against vendor quotations")

if not PDF_SUPPORT:
    st.warning("⚠️ PDF support unavailable. Install: `pip install pdfplumber`")

tabs = st.tabs(["📁 Upload & Configure", "📊 Dashboard", "🔍 Analysis", "📋 Summary", "💾 Export"])

# ===== TAB 1: Upload & Configure =====
with tabs[0]:
    # Drawing Upload
    st.subheader("1️⃣ Drawing/Schedule")
    draw_file = st.file_uploader("Upload drawing (PDF, Excel, CSV)", type=['pdf', 'csv', 'xlsx', 'xls'], key="draw")
    
    if draw_file and draw_file.name != st.session_state.drawing_filename:
        with st.spinner("Processing..."):
            dfs = parse_file(draw_file)
            if dfs:
                combined = max(dfs, key=len).reset_index(drop=True)
                st.session_state.drawing_df = combined
                st.session_state.drawing_filename = draw_file.name
                st.session_state.column_mapping = auto_detect_cols(combined, 'drawing')
                st.session_state.drawing_data = None
                st.rerun()
            else:
                st.error("Could not extract data")
    
    if st.session_state.drawing_filename:
        st.success(f"✅ {st.session_state.drawing_filename}")
    
    # Drawing Column Mapping
    if st.session_state.drawing_df is not None:
        st.markdown("**Map Drawing Columns:**")
        df = st.session_state.drawing_df
        opts = ['-- Not Used --'] + list(df.columns)
        
        with st.expander("Preview Drawing Data", expanded=False):
            st.dataframe(df.head(15), height=200, use_container_width=True)
        
        c1, c2, c3 = st.columns(3)
        mapping = st.session_state.column_mapping
        
        with c1:
            no_col = st.selectbox("Item No *", opts, index=opts.index(mapping.get('no')) if mapping.get('no') in opts else 0, key="d_no")
            desc_col = st.selectbox("Description *", opts, index=opts.index(mapping.get('description')) if mapping.get('description') in opts else 0, key="d_desc")
        with c2:
            qty_col = st.selectbox("Quantity", opts, index=opts.index(mapping.get('qty')) if mapping.get('qty') in opts else 0, key="d_qty")
            equip_col = st.selectbox("Equipment #", opts, index=opts.index(mapping.get('equip_num')) if mapping.get('equip_num') in opts else 0, key="d_equip")
        with c3:
            st.session_state.use_categories = st.checkbox("Use Categories", value=st.session_state.use_categories)
            cat_col = st.selectbox("Category", opts, index=opts.index(mapping.get('category')) if mapping.get('category') in opts else 0, key="d_cat") if st.session_state.use_categories else '-- Not Used --'
        
        if st.button("✅ Apply Drawing Mapping", type="primary"):
            new_map = {}
            for k, v in [('no', no_col), ('description', desc_col), ('qty', qty_col), ('equip_num', equip_col), ('category', cat_col)]:
                if v != '-- Not Used --':
                    new_map[k] = v
            st.session_state.column_mapping = new_map
            items = extract_drawing_data(df, new_map)
            if items:
                st.session_state.drawing_data = items
                st.success(f"✅ Extracted {len(items)} items")
                st.rerun()
            else:
                st.error("No items extracted. Check mapping.")
    
    if st.session_state.drawing_data:
        with st.expander(f"📋 Drawing Items ({len(st.session_state.drawing_data)})", expanded=False):
            st.dataframe(pd.DataFrame(st.session_state.drawing_data), height=250, use_container_width=True)
    
    # Quote Upload
    st.markdown("---")
    st.subheader("2️⃣ Quotations")
    quote_files = st.file_uploader("Upload quotes (PDF, Excel, CSV)", type=['pdf', 'csv', 'xlsx', 'xls'], accept_multiple_files=True, key="quotes")
    
    if quote_files:
        for qf in quote_files:
            if qf.name not in st.session_state.quote_dfs:
                with st.spinner(f"Processing {qf.name}..."):
                    dfs = parse_file(qf)
                    if dfs:
                        combined = pd.concat(dfs, ignore_index=True) if len(dfs) > 1 else dfs[0]
                        st.session_state.quote_dfs[qf.name] = combined.reset_index(drop=True)
                        st.session_state.quote_mappings[qf.name] = auto_detect_cols(combined, 'quote')
                        st.rerun()
                    else:
                        st.error(f"Could not extract from {qf.name}")
    
    # Quote Configuration
    if st.session_state.quote_dfs:
        for fname, qdf in st.session_state.quote_dfs.items():
            with st.expander(f"📄 {fname}", expanded=(fname not in st.session_state.quotes_data)):
                st.dataframe(qdf.head(15), height=180, use_container_width=True)
                
                opts = ['-- Not Used --'] + list(qdf.columns)
                qmap = st.session_state.quote_mappings.get(fname, {})
                
                qc1, qc2 = st.columns(2)
                with qc1:
                    q_no = st.selectbox("Item No", opts, index=opts.index(qmap.get('no')) if qmap.get('no') in opts else 0, key=f"qno_{fname}")
                    q_desc = st.selectbox("Description", opts, index=opts.index(qmap.get('description')) if qmap.get('description') in opts else 0, key=f"qdesc_{fname}")
                    q_qty = st.selectbox("Quantity", opts, index=opts.index(qmap.get('qty')) if qmap.get('qty') in opts else 0, key=f"qqty_{fname}")
                with qc2:
                    q_unit = st.selectbox("Unit Price (Sell)", opts, index=opts.index(qmap.get('unit_price')) if qmap.get('unit_price') in opts else 0, key=f"qunit_{fname}")
                    q_total = st.selectbox("Total Price (Sell Total)", opts, index=opts.index(qmap.get('total_price')) if qmap.get('total_price') in opts else 0, key=f"qtotal_{fname}")
                
                bc1, bc2 = st.columns(2)
                with bc1:
                    if st.button("✅ Apply", key=f"apply_{fname}", type="primary"):
                        new_qmap = {}
                        for k, v in [('no', q_no), ('description', q_desc), ('qty', q_qty), ('unit_price', q_unit), ('total_price', q_total)]:
                            if v != '-- Not Used --':
                                new_qmap[k] = v
                        st.session_state.quote_mappings[fname] = new_qmap
                        items = extract_quote_data(qdf, new_qmap, fname)
                        if items:
                            st.session_state.quotes_data[fname] = items
                            nic_count = sum(1 for i in items if i.get('Is_NIC'))
                            st.success(f"✅ {len(items)} items ({nic_count} NIC)")
                            st.rerun()
                        else:
                            st.error("No items extracted")
                with bc2:
                    if st.button("🗑️ Remove", key=f"rm_{fname}"):
                        del st.session_state.quote_dfs[fname]
                        st.session_state.quotes_data.pop(fname, None)
                        st.session_state.quote_mappings.pop(fname, None)
                        st.rerun()
                
                if fname in st.session_state.quotes_data:
                    items = st.session_state.quotes_data[fname]
                    nic_count = sum(1 for i in items if i.get('Is_NIC'))
                    total = sum(i['Total_Price'] for i in items if not i.get('Is_NIC'))
                    st.success(f"✅ {len(items)} items | {nic_count} NIC | ${total:,.2f}")
    
    # Summary & Reset
    if st.session_state.quotes_data:
        st.markdown("---")
        st.markdown("**Loaded Quotes:**")
        for fn, qs in st.session_state.quotes_data.items():
            nic = sum(1 for q in qs if q.get('Is_NIC'))
            total = sum(q['Total_Price'] for q in qs if not q.get('Is_NIC'))
            st.caption(f"• {fn}: {len(qs)} items ({nic} NIC) = ${total:,.2f}")
    
    if st.button("🔄 Reset All"):
        for key in ['drawing_data', 'drawing_df', 'quotes_data', 'quote_dfs', 'quote_mappings', 'drawing_filename', 'column_mapping']:
            st.session_state[key] = {} if 'data' in key or 'dfs' in key or 'mappings' in key else None
        st.rerun()

# ===== TAB 2: Dashboard =====
with tabs[1]:
    if not st.session_state.drawing_data:
        st.warning("⚠️ Upload and configure drawing first")
    elif not st.session_state.quotes_data:
        st.warning("⚠️ Upload and configure quotes first")
        if st.session_state.quote_dfs:
            st.info("💡 Click 'Apply' for each quote file to extract data")
    else:
        all_quotes = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_data(st.session_state.drawing_data, all_quotes, st.session_state.use_categories, st.session_state.supplier_codes)
        
        st.subheader("📊 Coverage Summary")
        
        # Actionable items
        exclude = ['Owner Supply', 'Existing', 'N/A']
        actionable = df[~df['Status'].isin(exclude)]
        
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("✅ Quoted", len(actionable[actionable['Status'] == 'Quoted']))
        c2.metric("❌ Missing", len(actionable[actionable['Status'] == 'MISSING']))
        c3.metric("⚠️ Qty Mismatch", len(actionable[actionable['Status'] == 'Qty Mismatch']))
        c4.metric("🚫 NIC", len(df[df['Status'] == 'NIC']))
        c5.metric("📋 Needs Pricing", len(actionable[actionable['Status'] == 'Needs Pricing']))
        
        col1, col2 = st.columns(2)
        col1.metric("💰 Total Quoted", f"${df['Total_Price'].sum():,.2f}")
        col2.metric("📦 Items", f"{len(df)} ({len(actionable)} actionable)")
        
        # Charts
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

# ===== TAB 3: Analysis =====
with tabs[2]:
    if st.session_state.drawing_data and st.session_state.quotes_data:
        all_quotes = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_data(st.session_state.drawing_data, all_quotes, st.session_state.use_categories, st.session_state.supplier_codes)
        
        st.subheader("🔍 Detailed Analysis")
        
        col1, col2 = st.columns(2)
        status_opts = df['Status'].unique().tolist()
        filt_status = col1.multiselect("Filter Status", status_opts, default=status_opts)
        
        fdf = df[df['Status'].isin(filt_status)]
        
        def highlight(row):
            colors = {'Quoted': 'background-color:#d4edda', 'MISSING': 'background-color:#f8d7da', 'Qty Mismatch': 'background-color:#fff3cd', 'NIC': 'background-color:#e2d5f0', 'Needs Pricing': 'background-color:#ffe5d0'}
            return [colors.get(row['Status'], '')] * len(row)
        
        cols = ['Drawing_No', 'Equip_Num', 'Description', 'Drawing_Qty']
        if st.session_state.use_categories:
            cols.append('Category')
        cols.extend(['Quote_Item_No', 'Quote_Qty', 'Unit_Price', 'Total_Price', 'Status', 'Issue'])
        
        st.dataframe(fdf[cols].style.apply(highlight, axis=1), height=400, use_container_width=True)
        
        # Critical missing
        st.subheader("🚨 Critical Missing")
        critical = df[(df['Status'] == 'MISSING') & (df['Category'].isin([5, 6]) if st.session_state.use_categories else True)]
        if len(critical) > 0:
            st.error(f"⚠️ {len(critical)} items need quotes!")
            st.dataframe(critical[['Drawing_No', 'Equip_Num', 'Description', 'Drawing_Qty']], use_container_width=True)
        else:
            st.success("✅ No critical missing items!")
        
        # NIC items
        nic_items = df[df['Status'] == 'NIC']
        if len(nic_items) > 0:
            st.subheader("🚫 NIC Items (Not In Contract)")
            st.dataframe(nic_items[['Drawing_No', 'Description', 'Drawing_Qty']], use_container_width=True)
    else:
        st.warning("⚠️ Configure drawing and quotes first")

# ===== TAB 4: Summary =====
with tabs[3]:
    if st.session_state.drawing_data and st.session_state.quotes_data:
        all_quotes = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_data(st.session_state.drawing_data, all_quotes, st.session_state.use_categories, st.session_state.supplier_codes)
        
        st.subheader("📋 Summary by Status")
        summary = df.groupby('Status').agg(Items=('Status', 'count'), Total_Value=('Total_Price', 'sum')).reset_index()
        summary['Total_Value'] = summary['Total_Value'].apply(lambda x: f"${x:,.2f}")
        st.dataframe(summary, use_container_width=True)
        
        st.subheader("📄 Quote Files")
        for fn, qs in st.session_state.quotes_data.items():
            nic = sum(1 for q in qs if q.get('Is_NIC'))
            total = sum(q['Total_Price'] for q in qs if not q.get('Is_NIC'))
            st.caption(f"• {fn}: {len(qs)} items ({nic} NIC) = ${total:,.2f}")
    else:
        st.warning("⚠️ Configure drawing and quotes first")

# ===== TAB 5: Export =====
with tabs[4]:
    st.subheader("💾 Export")
    
    if st.session_state.drawing_data and st.session_state.quotes_data:
        all_quotes = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_data(st.session_state.drawing_data, all_quotes, st.session_state.use_categories, st.session_state.supplier_codes)
        
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Full Analysis', index=False)
            df[df['Status'] == 'MISSING'].to_excel(writer, sheet_name='Missing', index=False)
            df[df['Status'] == 'Quoted'].to_excel(writer, sheet_name='Quoted', index=False)
            df[df['Status'] == 'NIC'].to_excel(writer, sheet_name='NIC Items', index=False)
            pd.DataFrame(all_quotes).to_excel(writer, sheet_name='All Quotes', index=False)
        out.seek(0)
        
        st.download_button("📥 Download Full Report", out, f"Analysis_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    else:
        st.warning("⚠️ Configure drawing and quotes first")

st.markdown("---")
st.caption("Drawing Quote Analyzer v8.0 | NIC = Not In Contract")
