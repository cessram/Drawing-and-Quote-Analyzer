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

# Default supplier codes (can be customized)
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
        # Convert to string and clean
        col_name = str(c).strip() if pd.notna(c) and str(c).strip() != '' else f'Column_{i}'
        
        # Handle duplicates
        if col_name in seen:
            seen[col_name] += 1
            col_name = f"{col_name}_{seen[col_name]}"
        else:
            seen[col_name] = 0
        
        new_cols.append(col_name)
    
    df.columns = new_cols
    df = df.dropna(how='all')
    return df

def parse_pdf_to_text(uploaded_file):
    """Extract text from PDF."""
    if not PDF_SUPPORT:
        return None
    text_content = []
    uploaded_file.seek(0)
    try:
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    text_content.append(text)
    except Exception as e:
        st.warning(f"Could not extract text from PDF: {e}")
        return None
    return "\n".join(text_content)

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
                        # Create DataFrame with first row as header
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
        tables = parse_pdf_tables(uploaded_file)
        return tables
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
            'no': ['no', 'no.', 'item', 'item #', 'item no', 'number', '#', 'id', 'ref', 'line'],
            'description': ['description', 'desc', 'equipment', 'name', 'item description', 'material', 'product'],
            'qty': ['qty', 'qty.', 'quantity', 'count', 'amount', 'units'],
            'unit_price': ['unit price', 'price', 'sell', 'rate', 'unit cost', 'cost ea', 'each'],
            'total_price': ['total', 'total price', 'sell total', 'ext price', 'extended', 'amount', 'line total'],
            'unit': ['unit', 'uom', 'measure']
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
    """Extract quote data using column mapping."""
    items = []
    
    no_col = col_map.get('no')
    desc_col = col_map.get('description')
    qty_col = col_map.get('qty')
    unit_col = col_map.get('unit_price')
    total_col = col_map.get('total_price')
    
    for idx, row in df.iterrows():
        try:
            desc_val = str(row.get(desc_col, '')).strip() if desc_col else ''
            
            # Skip empty/header rows
            if desc_col and (not desc_val or desc_val.lower() in ('nan', '', 'description', 'none', 'nic')):
                continue
            
            no_val = ''
            if no_col:
                no_val = str(row.get(no_col, '')).strip()
                if no_val.lower() in ('nan', 'none'):
                    no_val = ''
            
            qty = 1
            if qty_col:
                qty_val = clean_numeric(row.get(qty_col, 1))
                if qty_val:
                    qty = int(qty_val)
            
            unit_price = 0
            if unit_col:
                up = clean_numeric(row.get(unit_col, 0))
                if up:
                    unit_price = up
            
            total_price = 0
            if total_col:
                tp = clean_numeric(row.get(total_col, 0))
                if tp:
                    total_price = tp
            
            if total_price == 0 and unit_price > 0:
                total_price = unit_price * qty
            
            items.append({
                'Item_No': no_val,
                'Description': desc_val,
                'Qty': qty,
                'Unit_Price': unit_price,
                'Total_Price': total_price,
                'Source_File': source_file
            })
        except:
            continue
    
    return items

def match_items(drawing_no, quotes):
    """Match drawing item to quote by item number."""
    drawing_no_clean = str(drawing_no).strip().lower()
    
    # Exact match
    for q in quotes:
        if str(q.get('Item_No', '')).strip().lower() == drawing_no_clean:
            return q
    
    # Numeric match (ignore letters)
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
            if match['Qty'] == item['Qty']:
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
            'Total_Price': match['Total_Price'] if match else 0,
            'Quote_Source': match['Source_File'] if match else '-',
            'Status': status,
            'Issue': issue
        })
    
    return pd.DataFrame(analysis)

# ===== UI =====
st.markdown("## 📊 Drawing vs Quote Analyzer")
st.caption("Universal tool for comparing drawings/schedules against vendor quotations")

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
                        # Select the largest DataFrame (most likely main data)
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
        
        # Apply mapping button
        if st.button("✅ Apply Column Mapping", type="primary"):
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
            
            # Extract data
            items = extract_drawing_data(df, mapping)
            if items:
                st.session_state.drawing_data = items
                st.success(f"✅ Extracted {len(items)} items from drawing")
                st.rerun()
            else:
                st.error("Could not extract data. Check column mapping.")
    
    # Show extracted drawing data
    if st.session_state.drawing_data:
        with st.expander(f"📋 Extracted Drawing Items ({len(st.session_state.drawing_data)} items)", expanded=False):
            st.dataframe(pd.DataFrame(st.session_state.drawing_data), height=300, use_container_width=True)
    
    st.markdown("---")
    st.subheader("3️⃣ Upload Quotations")
    
    qcol1, qcol2 = st.columns([2, 1])
    
    with qcol1:
        quote_files = st.file_uploader(
            "Upload quote files (PDF, Excel, CSV)", 
            type=['pdf', 'csv', 'xlsx', 'xls'], 
            accept_multiple_files=True,
            key="quote_upload"
        )
        
        if quote_files:
            for qf in quote_files:
                if qf.name not in st.session_state.quotes_data:
                    with st.spinner(f"Processing {qf.name}..."):
                        dfs = parse_uploaded_file(qf)
                        if dfs and len(dfs) > 0:
                            all_items = []
                            for qdf in dfs:
                                col_map = auto_detect_columns(qdf, 'quote')
                                items = extract_quote_data(qdf, col_map, qf.name)
                                all_items.extend(items)
                            
                            if all_items:
                                st.session_state.quotes_data[qf.name] = all_items
                                st.success(f"✅ {len(all_items)} items from {qf.name}")
    
    with qcol2:
        if st.session_state.quotes_data:
            st.success(f"📄 {len(st.session_state.quotes_data)} quote file(s)")
            for fn, qs in st.session_state.quotes_data.items():
                total = sum(q['Total_Price'] for q in qs)
                st.caption(f"• {fn}: {len(qs)} items (${total:,.2f})")
            
            if st.button("🗑️ Clear All Quotes"):
                st.session_state.quotes_data = {}
                st.rerun()
    
    st.markdown("---")
    if st.button("🔄 Reset Everything"):
        st.session_state.drawing_data = None
        st.session_state.drawing_df = None
        st.session_state.quotes_data = {}
        st.session_state.drawing_filename = None
        st.session_state.column_mapping = {}
        st.rerun()

# ===== TAB 2: Dashboard =====
with tabs[1]:
    if not st.session_state.drawing_data:
        st.warning("⚠️ Please upload and configure drawing first")
    elif not st.session_state.quotes_data:
        st.warning("⚠️ Please upload quotations")
    else:
        all_quotes = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_data(
            st.session_state.drawing_data, 
            all_quotes, 
            st.session_state.use_categories,
            st.session_state.supplier_codes
        )
        
        st.subheader("📊 Coverage Summary")
        
        # Actionable items (exclude owner supply, existing, N/A)
        if st.session_state.use_categories:
            actionable = df[~df['Status'].isin(['Owner Supply', 'Existing', 'N/A'])]
        else:
            actionable = df[~df['Status'].isin(['N/A'])]
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("✅ Quoted", len(actionable[actionable['Status'] == 'Quoted']))
        c2.metric("❌ Missing", len(actionable[actionable['Status'] == 'MISSING']))
        c3.metric("⚠️ Qty Mismatch", len(actionable[actionable['Status'] == 'Qty Mismatch']))
        c4.metric("📋 Needs Pricing", len(actionable[actionable['Status'] == 'Needs Pricing']))
        
        col1, col2 = st.columns(2)
        col1.metric("💰 Total Quoted Value", f"${df['Total_Price'].sum():,.2f}")
        col2.metric("📦 Total Items", f"{len(df)} ({len(actionable)} actionable)")
        
        ch1, ch2 = st.columns(2)
        
        with ch1:
            vc = df['Status'].value_counts().reset_index()
            vc.columns = ['Status', 'Count']
            colors = {
                'Quoted': '#28a745', 'MISSING': '#dc3545', 'Qty Mismatch': '#ffc107',
                'Needs Pricing': '#fd7e14', 'Owner Supply': '#6c757d', 
                'Existing': '#adb5bd', 'N/A': '#e9ecef'
            }
            fig = px.pie(vc, values='Count', names='Status', color='Status', 
                        color_discrete_map=colors, title="Status Distribution")
            st.plotly_chart(fig, use_container_width=True)
        
        with ch2:
            if st.session_state.use_categories and df['Category'].notna().any():
                cat_df = df[df['Category'].notna()].groupby('Category').size().reset_index(name='Items')
                cat_df['Label'] = cat_df['Category'].astype(int).apply(lambda x: f"Cat {x}")
                fig2 = px.bar(cat_df, x='Label', y='Items', title="Items by Category")
                st.plotly_chart(fig2, use_container_width=True)
            else:
                # Show status breakdown as bar
                fig2 = px.bar(vc, x='Status', y='Count', title="Items by Status", color='Status',
                             color_discrete_map=colors)
                st.plotly_chart(fig2, use_container_width=True)

# ===== TAB 3: Analysis =====
with tabs[2]:
    if st.session_state.drawing_data and st.session_state.quotes_data:
        all_quotes = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_data(
            st.session_state.drawing_data, 
            all_quotes,
            st.session_state.use_categories,
            st.session_state.supplier_codes
        )
        
        st.subheader("🔍 Detailed Analysis")
        
        col1, col2 = st.columns(2)
        status_opts = df['Status'].unique().tolist()
        filt_status = col1.multiselect("Filter by Status", status_opts, default=status_opts)
        
        if st.session_state.use_categories and df['Category'].notna().any():
            cat_opts = sorted([int(c) for c in df['Category'].dropna().unique()])
            filt_cat = col2.multiselect("Filter by Category", cat_opts, default=cat_opts)
        else:
            filt_cat = None
        
        # Filter data
        fdf = df[df['Status'].isin(filt_status)]
        if filt_cat:
            fdf = fdf[(fdf['Category'].isin(filt_cat)) | (fdf['Category'].isna())]
        
        # Highlight function
        def highlight_row(row):
            colors = {
                'Quoted': 'background-color:#d4edda',
                'MISSING': 'background-color:#f8d7da',
                'Qty Mismatch': 'background-color:#fff3cd',
                'Needs Pricing': 'background-color:#ffe5d0'
            }
            return [colors.get(row['Status'], '')] * len(row)
        
        # Select columns to display
        display_cols = ['Drawing_No', 'Equip_Num', 'Description', 'Drawing_Qty']
        if st.session_state.use_categories:
            display_cols.append('Category')
        display_cols.extend(['Quote_Item_No', 'Quote_Qty', 'Unit_Price', 'Total_Price', 'Status', 'Issue'])
        
        st.dataframe(
            fdf[display_cols].style.apply(highlight_row, axis=1),
            height=450, 
            use_container_width=True
        )
        
        # Critical missing section
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

# ===== TAB 4: Summary =====
with tabs[3]:
    if st.session_state.drawing_data and st.session_state.quotes_data:
        all_quotes = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_data(
            st.session_state.drawing_data, 
            all_quotes,
            st.session_state.use_categories,
            st.session_state.supplier_codes
        )
        
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
                        'Value': f"${ci['Total_Price'].sum():,.2f}"
                    })
            if summary:
                st.dataframe(pd.DataFrame(summary), use_container_width=True)
        
        st.subheader("📋 Summary by Status")
        status_summary = df.groupby('Status').agg(
            Items=('Status', 'count'),
            Total_Value=('Total_Price', 'sum')
        ).reset_index()
        status_summary['Total_Value'] = status_summary['Total_Value'].apply(lambda x: f"${x:,.2f}")
        st.dataframe(status_summary, use_container_width=True)
        
        # Quote file summary
        st.subheader("📄 Quote Files Summary")
        quote_summary = []
        for fn, qs in st.session_state.quotes_data.items():
            quote_summary.append({
                'File': fn,
                'Items': len(qs),
                'Total Value': f"${sum(q['Total_Price'] for q in qs):,.2f}"
            })
        st.dataframe(pd.DataFrame(quote_summary), use_container_width=True)
        
        # Items without category
        no_cat = df[df['Category'].isna()]
        if len(no_cat) > 0:
            st.subheader("📌 Items Without Category / SPARE")
            st.dataframe(no_cat[['Drawing_No', 'Equip_Num', 'Description', 'Drawing_Qty', 'Status']], use_container_width=True)

# ===== TAB 5: Export =====
with tabs[4]:
    st.subheader("💾 Export Data")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.session_state.drawing_data:
            st.markdown("**Drawing Data**")
            eq_df = pd.DataFrame(st.session_state.drawing_data)
            out_eq = io.BytesIO()
            eq_df.to_excel(out_eq, index=False)
            out_eq.seek(0)
            st.download_button(
                "📥 Download Drawing Data",
                out_eq,
                f"Drawing_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    with col2:
        if st.session_state.quotes_data:
            st.markdown("**Quote Data**")
            all_q = [q for qs in st.session_state.quotes_data.values() for q in qs]
            q_df = pd.DataFrame(all_q)
            out_q = io.BytesIO()
            q_df.to_excel(out_q, index=False)
            out_q.seek(0)
            st.download_button(
                "📥 Download Quote Data",
                out_q,
                f"Quotes_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    st.markdown("---")
    
    if st.session_state.drawing_data and st.session_state.quotes_data:
        st.markdown("**Full Analysis Report**")
        
        all_quotes = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_data(
            st.session_state.drawing_data, 
            all_quotes,
            st.session_state.use_categories,
            st.session_state.supplier_codes
        )
        
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Full Analysis', index=False)
            df[df['Status'] == 'MISSING'].to_excel(writer, sheet_name='Missing Items', index=False)
            df[df['Status'] == 'Quoted'].to_excel(writer, sheet_name='Quoted Items', index=False)
            df[df['Status'] == 'Qty Mismatch'].to_excel(writer, sheet_name='Qty Mismatch', index=False)
            pd.DataFrame(all_quotes).to_excel(writer, sheet_name='All Quotes', index=False)
        out.seek(0)
        
        st.download_button(
            "📥 Download Full Analysis Report",
            out,
            f"Analysis_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

st.markdown("---")
st.caption("Universal Drawing Quote Analyzer v7.2")
