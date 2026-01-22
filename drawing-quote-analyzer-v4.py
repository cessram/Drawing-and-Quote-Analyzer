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

st.set_page_config(page_title="Drawing Quote Analyzer", page_icon="🔍", layout="wide")

SUPPLIER_CODES = {
    1: "IH SUPPLY / IH INSTALL",
    2: "IH SUPPLY / IH INSTALL (DH EQUIPMENT)",
    3: "IH SUPPLY / IH INSTALL (BIO MED)",
    4: "IH SUPPLY / VENDOR INSTALL",
    5: "CONTRACTOR SUPPLY / CONTRACTOR INSTALL",
    6: "CONTRACTOR SUPPLY / VENDOR INSTALL",
    7: "IH SUPPLY / CONTRACTOR INSTALL",
    8: "EXISTING / RELOCATED EQUIPMENT"
}

if 'equipment_schedule' not in st.session_state:
    st.session_state.equipment_schedule = None
if 'quotes_data' not in st.session_state:
    st.session_state.quotes_data = {}
if 'drawing_filename' not in st.session_state:
    st.session_state.drawing_filename = None

def parse_pdf_file(uploaded_file):
    if not PDF_SUPPORT:
        st.error("PDF support not available. Install pdfplumber")
        return None, None
    text_content = []
    all_tables = []
    uploaded_file.seek(0)
    with pdfplumber.open(uploaded_file) as pdf:
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text()
            if text:
                text_content.append(text)
            tables = page.extract_tables()
            for table in tables:
                if table and len(table) > 1:
                    try:
                        df = pd.DataFrame(table[1:], columns=table[0] if table[0] else None)
                        all_tables.append(df)
                    except:
                        pass
    return "\n".join(text_content), all_tables

def parse_excel_file(uploaded_file):
    try:
        uploaded_file.seek(0)
        xl = pd.ExcelFile(uploaded_file)
        return {name: pd.read_excel(xl, sheet_name=name) for name in xl.sheet_names}
    except Exception as e:
        st.error(f"Error reading Excel: {e}")
        return None

def parse_csv_file(uploaded_file):
    try:
        uploaded_file.seek(0)
        return {"Sheet1": pd.read_csv(uploaded_file)}
    except Exception as e:
        st.error(f"Error reading CSV: {e}")
        return None

def parse_uploaded_file(uploaded_file):
    ext = uploaded_file.name.split('.')[-1].lower()
    if ext == 'pdf':
        text, tables = parse_pdf_file(uploaded_file)
        return {'type': 'pdf', 'text': text, 'tables': tables, 'filename': uploaded_file.name}
    elif ext in ['xlsx', 'xls']:
        sheets = parse_excel_file(uploaded_file)
        return {'type': 'excel', 'sheets': sheets, 'filename': uploaded_file.name}
    elif ext == 'csv':
        sheets = parse_csv_file(uploaded_file)
        return {'type': 'csv', 'sheets': sheets, 'filename': uploaded_file.name}
    return None

def show_file_preview(parsed_data):
    """Show a preview of the file structure to help with debugging"""
    if parsed_data['type'] == 'pdf':
        if parsed_data['tables']:
            st.write(f"**PDF has {len(parsed_data['tables'])} table(s)**")
            for i, tbl in enumerate(parsed_data['tables']):
                st.write(f"Table {i+1} columns: {tbl.columns.tolist()}")
                st.dataframe(tbl.head(5), height=150)
        else:
            st.warning("No tables found in PDF")
    elif parsed_data['type'] in ['excel', 'csv']:
        if parsed_data.get('sheets'):
            for name, df in parsed_data['sheets'].items():
                st.write(f"**Sheet '{name}'** - Columns: {df.columns.tolist()}")
                st.dataframe(df.head(5), height=150)

def extract_equipment_from_dataframe(df, debug=False):
    """Extract equipment items from a dataframe with flexible column detection"""
    equipment_list = []
    original_cols = df.columns.tolist()
    
    # Clean up the dataframe - remove completely empty rows
    df = df.dropna(how='all').reset_index(drop=True)
    
    # Normalize column names
    df.columns = df.columns.astype(str).str.strip().str.lower()
    
    if debug:
        st.write("**Original columns:**", original_cols)
        st.write("**Normalized columns:**", df.columns.tolist())
        st.write("**DataFrame shape:**", df.shape)
    
    # Extended column mapping with more variations
    col_map = {
        'no': ['no', 'no.', 'item', 'item #', 'item no', 'item no.', 'number', '#', 
               'eq no', 'eq no.', 'equipment no', 'equipment no.', 'equip no', 
               'equip no.', 'id', 'ref', 'ref.', 'reference', 'tag', 'tag no', 'tag no.'],
        'description': ['description', 'desc', 'equipment', 'name', 'item description', 
                       'equipment description', 'equip desc', 'equipment name', 
                       'item name', 'remarks', 'details'],
        'qty': ['qty', 'quantity', 'count', 'qnty', 'qty.', 'amount', 'units', 'ea'],
        'category': ['category', 'cat', 'supplier code', 'code', 'supplier', 'type', 
                    'supply', 'source', 'cat.', 'supplier cat']
    }
    
    found = {}
    
    # First pass: exact match
    for key, opts in col_map.items():
        for col in df.columns:
            col_clean = col.lower().strip()
            if col_clean in opts:
                found[key] = col
                break
    
    # Second pass: partial match for unfound columns
    for key, opts in col_map.items():
        if key not in found:
            for col in df.columns:
                col_clean = col.lower().strip()
                if any(o in col_clean for o in opts):
                    found[key] = col
                    break
    
    if debug:
        st.write("**Found column mapping:**", found)
    
    # Fallback: try using first columns if standard mapping failed
    if 'no' not in found or 'description' not in found:
        if len(df.columns) >= 2:
            first_col = df.columns[0]
            second_col = df.columns[1]
            
            if debug:
                st.write(f"**Trying fallback:** First col '{first_col}' as No, Second col '{second_col}' as Description")
            
            # Check if first column looks like item numbers (e.g., "1", "1a", "23", etc.)
            sample_vals = df[first_col].dropna().head(10).astype(str).tolist()
            looks_like_numbers = any(
                re.match(r'^\d+[a-zA-Z]?$', str(v).strip()) 
                for v in sample_vals if str(v).strip()
            )
            
            if looks_like_numbers:
                if 'no' not in found:
                    found['no'] = first_col
                if 'description' not in found:
                    found['description'] = second_col
                if debug:
                    st.write("**Fallback accepted** - first column contains item numbers")
            elif debug:
                st.write("**Fallback rejected** - first column doesn't look like item numbers")
                st.write("Sample values:", sample_vals[:5])
    
    # Additional fallback: check for any column that looks like it contains item numbers
    if 'no' not in found:
        for col in df.columns:
            sample = df[col].dropna().head(10).astype(str).tolist()
            if any(re.match(r'^\d+[a-zA-Z]?$', str(v).strip()) for v in sample if str(v).strip()):
                found['no'] = col
                if debug:
                    st.write(f"**Auto-detected 'no' column:** {col}")
                break
    
    # Check for description in remaining columns
    if 'description' not in found and 'no' in found:
        for col in df.columns:
            if col != found['no']:
                # Check if column has text content (longer strings)
                sample = df[col].dropna().head(10).astype(str).tolist()
                avg_len = sum(len(str(v)) for v in sample) / max(len(sample), 1)
                if avg_len > 10:  # Descriptions are usually longer
                    found['description'] = col
                    if debug:
                        st.write(f"**Auto-detected 'description' column:** {col}")
                    break
    
    if 'no' not in found or 'description' not in found:
        if debug:
            st.error(f"Missing required columns. Found mapping: {found}")
            st.info("Need columns for: Item Number (no) and Description")
        return None
    
    # Extract equipment items
    for idx, row in df.iterrows():
        try:
            no = str(row.get(found.get('no', ''), '')).strip()
            desc = str(row.get(found.get('description', ''), '')).strip()
            
            # Skip empty/invalid rows
            if not no or no.lower() in ['nan', '', 'none', 'no', 'no.', 'item', 'item no', 'item no.']:
                continue
            if not desc or desc.lower() in ['nan', '', 'none', 'description', 'desc']:
                continue
            
            # Skip header-like rows
            if no.lower() == found.get('no', '').lower():
                continue
            
            qty = 1
            if 'qty' in found:
                try:
                    qval = str(row.get(found['qty'], 1)).replace(',', '').strip()
                    qval = re.sub(r'[^\d.]', '', qval)  # Remove non-numeric chars
                    qty = int(float(qval)) if qval and qval.lower() not in ['nan', ''] else 1
                except:
                    pass
            
            cat = None
            if 'category' in found:
                try:
                    cval = str(row.get(found['category'], '')).strip()
                    cval = re.sub(r'[^\d]', '', cval)  # Extract only digits
                    cat = int(float(cval)) if cval else None
                except:
                    pass
            
            equipment_list.append({'No': no, 'Description': desc, 'Qty': qty, 'Category': cat})
        except Exception as e:
            if debug:
                st.write(f"Row {idx} error: {e}")
            continue
    
    if debug:
        st.write(f"**Extracted {len(equipment_list)} items**")
        if equipment_list:
            st.write("**Sample items:**")
            for item in equipment_list[:3]:
                st.write(f"  - {item}")
    
    return equipment_list if equipment_list else None

def process_drawing_file(parsed_data, debug=False):
    """Process drawing file and extract equipment list"""
    equipment_list = []
    
    if parsed_data['type'] == 'pdf' and parsed_data['tables']:
        if debug:
            st.write(f"**PDF Tables found:** {len(parsed_data['tables'])}")
        for i, tbl in enumerate(parsed_data['tables']):
            if debug:
                st.write(f"**Processing Table {i+1}:**")
            ext = extract_equipment_from_dataframe(tbl, debug=debug)
            if ext:
                equipment_list.extend(ext)
                
    elif parsed_data['type'] in ['excel', 'csv'] and parsed_data.get('sheets'):
        if debug:
            st.write(f"**Sheets found:** {list(parsed_data['sheets'].keys())}")
        for name, df in parsed_data['sheets'].items():
            if debug:
                st.write(f"**Processing Sheet '{name}':**")
            ext = extract_equipment_from_dataframe(df, debug=debug)
            if ext:
                equipment_list.extend(ext)
    
    if debug and not equipment_list:
        st.warning("No equipment extracted from any table/sheet")
    
    # Remove duplicates
    seen = set()
    unique = []
    for item in equipment_list:
        key = (item['No'], item['Description'])
        if key not in seen:
            seen.add(key)
            unique.append(item)
    
    return unique

def parse_crs_quote_from_text(text, filename):
    quotes = []
    lines = text.split('\n')
    current_item = None
    current_qty = 0
    current_desc = ""
    current_unit = 0
    current_total = 0
    
    skip_words = ['page ', 'canadian restaurant', 'bird construc', 'fwg ltc', 
                  'item qty description', 'sell total', 'merchandise', 
                  'prices are in', 'quote valid']
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        ll = line.lower()
        if any(s in ll for s in skip_words):
            continue
        if re.match(r'^[\d\-]+[a-z]?\s+NIC', line, re.IGNORECASE):
            continue
        
        total_m = re.search(r'ITEM\s*TOTAL[:\s]*\$?([\d,]+\.\d{2})', line, re.IGNORECASE)
        if total_m and current_item:
            current_total = float(total_m.group(1).replace(',', ''))
            if not current_unit and current_qty > 0:
                current_unit = current_total / current_qty
            continue
        
        m = re.match(r'^(\d+[a-z]?)\s+(\d+)\s*ea\s+(.+)$', line, re.IGNORECASE)
        if m:
            if current_item and current_desc:
                quotes.append({
                    'Item': current_item,
                    'Description': current_desc.strip(),
                    'Qty': current_qty,
                    'Unit_Price': current_unit,
                    'Total_Price': current_total if current_total else current_unit * current_qty,
                    'Source_File': filename
                })
            current_item = m.group(1)
            current_qty = int(m.group(2))
            rest = m.group(3).strip()
            prices = re.findall(r'\$?([\d,]+\.\d{2})', rest)
            current_desc = re.sub(r'\s*\$?[\d,]+\.\d{2}', '', rest).strip()
            current_unit = 0
            current_total = 0
            if len(prices) >= 2:
                current_unit = float(prices[0].replace(',', ''))
                current_total = float(prices[1].replace(',', ''))
            elif len(prices) == 1:
                current_unit = float(prices[0].replace(',', ''))
                current_total = current_unit * current_qty
            continue
    
    if current_item and current_desc:
        quotes.append({
            'Item': current_item,
            'Description': current_desc.strip(),
            'Qty': current_qty,
            'Unit_Price': current_unit,
            'Total_Price': current_total if current_total else current_unit * current_qty,
            'Source_File': filename
        })
    return quotes

def extract_quotes_from_dataframe(df, filename):
    quotes = []
    df.columns = df.columns.astype(str).str.strip().str.lower()
    col_map = {
        'item': ['item', 'item no', 'no', 'no.', 'number'],
        'description': ['description', 'desc', 'equipment', 'name'],
        'qty': ['qty', 'quantity', 'count'],
        'unit_price': ['sell', 'unit price', 'price', 'each'],
        'total_price': ['sell total', 'total', 'total price', 'amount']
    }
    found = {}
    for key, opts in col_map.items():
        for col in df.columns:
            if any(o == col.lower().strip() for o in opts):
                found[key] = col
                break
    for _, row in df.iterrows():
        try:
            desc = str(row.get(found.get('description', ''), '')).strip()
            if not desc or desc.lower() in ['nan', '', 'nic']:
                continue
            item = str(row.get(found.get('item', ''), '')).strip()
            qty = 1
            try:
                qty = int(float(str(row.get(found.get('qty', ''), 1)).replace('ea', '').replace(',', '').strip()))
            except:
                pass
            up = 0
            try:
                up = float(str(row.get(found.get('unit_price', ''), 0)).replace('$', '').replace(',', '').strip())
            except:
                pass
            tp = 0
            try:
                tp = float(str(row.get(found.get('total_price', ''), 0)).replace('$', '').replace(',', '').strip())
            except:
                tp = up * qty
            quotes.append({'Item': item, 'Description': desc, 'Qty': qty, 'Unit_Price': up, 'Total_Price': tp, 'Source_File': filename})
        except:
            continue
    return quotes

def process_quote_file(parsed_data):
    quotes = []
    if parsed_data['type'] == 'pdf':
        if parsed_data['text']:
            quotes = parse_crs_quote_from_text(parsed_data['text'], parsed_data['filename'])
        if not quotes and parsed_data['tables']:
            for tbl in parsed_data['tables']:
                ext = extract_quotes_from_dataframe(tbl, parsed_data['filename'])
                if ext:
                    quotes.extend(ext)
    elif parsed_data['type'] in ['excel', 'csv'] and parsed_data.get('sheets'):
        for df in parsed_data['sheets'].values():
            ext = extract_quotes_from_dataframe(df, parsed_data['filename'])
            if ext:
                quotes.extend(ext)
    seen = set()
    unique = []
    for q in quotes:
        if q['Item'] and q['Item'] not in seen:
            seen.add(q['Item'])
            unique.append(q)
    return unique

def match_quote_to_schedule(item, quotes):
    no = str(item['No']).strip().lower()
    for q in quotes:
        qi = str(q.get('Item', '')).strip().lower()
        if no == qi:
            return q
    try:
        no_int = int(re.sub(r'[a-zA-Z]', '', no))
        for q in quotes:
            try:
                qi_int = int(re.sub(r'[a-zA-Z]', '', str(q.get('Item', '')).strip()))
                if no_int == qi_int:
                    return q
            except:
                pass
    except:
        pass
    return None

def analyze_schedule_vs_quotes(schedule, quotes):
    analysis = []
    for item in schedule:
        match = match_quote_to_schedule(item, quotes)
        cat = item.get('Category')
        if cat in [1, 2, 3]:
            status, issue = "IH Supply", "IH handles supply & install"
        elif cat == 8:
            status, issue = "Existing", "Existing/relocated"
        elif cat is None:
            status, issue = "N/A", "Spare or placeholder"
        elif match:
            if match['Qty'] == item['Qty']:
                status, issue = "✓ Quoted", None
            elif match['Qty'] > 0:
                status, issue = "⚠ Qty Mismatch", f"Expected {item['Qty']}, got {match['Qty']}"
            else:
                status, issue = "⚡ Included", "Part of system"
        else:
            if cat == 7:
                status, issue = "⚠ Needs Install", "IH supplies - needs install pricing"
            elif cat in [5, 6]:
                status, issue = "❌ MISSING", "Critical - requires quote"
            else:
                status, issue = "❌ MISSING", "Not found"
        analysis.append({
            'No': item['No'],
            'Quote_Item': match['Item'] if match else '-',
            'Description': item['Description'],
            'Schedule_Qty': item['Qty'],
            'Quote_Qty': match['Qty'] if match else 0,
            'Supplier_Code': cat,
            'Supplier_Desc': SUPPLIER_CODES.get(cat, 'Unknown'),
            'Unit_Price': match['Unit_Price'] if match else 0,
            'Total_Price': match['Total_Price'] if match else 0,
            'Source_File': match['Source_File'] if match else '-',
            'Status': status,
            'Issue': issue
        })
    return pd.DataFrame(analysis)

# ==================== UI ====================
st.markdown("## 🔍 Equipment Quote Analyzer")
st.markdown("Upload drawings and quotations for analysis")

if not PDF_SUPPORT:
    st.warning("PDF support not installed. Run: pip install pdfplumber")

tabs = st.tabs(["📤 Upload", "📊 Dashboard", "📋 Report", "🔢 Summary", "📥 Export"])

with tabs[0]:
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("📐 Drawing / Equipment Schedule")
        
        # Debug mode toggle
        debug_mode = st.checkbox("🔧 Debug Mode (show column detection)", key="debug_draw")
        
        # Show current status
        if st.session_state.equipment_schedule and len(st.session_state.equipment_schedule) > 0:
            st.success(f"✅ Loaded: {st.session_state.drawing_filename} ({len(st.session_state.equipment_schedule)} items)")
            with st.expander("View Equipment Schedule"):
                st.dataframe(pd.DataFrame(st.session_state.equipment_schedule), height=300)
        
        df_file = st.file_uploader("Select Drawing File", type=['pdf', 'csv', 'xlsx', 'xls'], key="draw")
        if df_file:
            with st.spinner("Processing drawing..."):
                parsed = parse_uploaded_file(df_file)
                if parsed:
                    # Always show preview in debug mode
                    if debug_mode:
                        st.write(f"**File type:** {parsed['type']}")
                        with st.expander("📋 Raw File Preview", expanded=True):
                            show_file_preview(parsed)
                    
                    equip = process_drawing_file(parsed, debug=debug_mode)
                    if equip and len(equip) > 0:
                        st.session_state.equipment_schedule = equip
                        st.session_state.drawing_filename = df_file.name
                        st.success(f"✅ Extracted {len(equip)} items from {df_file.name}")
                        if not debug_mode:
                            st.rerun()
                    else:
                        st.error("❌ Could not extract equipment.")
                        st.info("""
**Tips to fix this:**
1. Enable **Debug Mode** above to see what columns were detected
2. Your file needs columns similar to: `No.` or `Item` AND `Description`
3. Check that your data starts from row 1 (or has a header row)
4. Try exporting your drawing schedule to CSV/Excel format
                        """)
                        # Show preview anyway to help debug
                        if not debug_mode:
                            with st.expander("📋 Click to see file structure"):
                                show_file_preview(parsed)
                else:
                    st.error("❌ Could not parse file.")
    
    with c2:
        st.subheader("📝 Quotations")
        
        # Show current status
        if st.session_state.quotes_data and len(st.session_state.quotes_data) > 0:
            st.success(f"✅ {len(st.session_state.quotes_data)} quote file(s) loaded")
            for fn, qs in st.session_state.quotes_data.items():
                st.markdown(f"- **{fn}**: {len(qs)} items (${sum(q['Total_Price'] for q in qs):,.2f})")
        
        qf = st.file_uploader("Select Quote Files", type=['pdf', 'csv', 'xlsx', 'xls'], key="quote", accept_multiple_files=True)
        if qf:
            for f in qf:
                if f.name not in st.session_state.quotes_data:
                    with st.spinner(f"Processing {f.name}..."):
                        parsed = parse_uploaded_file(f)
                        if parsed:
                            q = process_quote_file(parsed)
                            if q and len(q) > 0:
                                st.session_state.quotes_data[f.name] = q
                                st.success(f"✅ Extracted {len(q)} items from {f.name}")
                                with st.expander(f"Preview: {f.name}"):
                                    st.dataframe(pd.DataFrame(q), height=200)
                            else:
                                st.warning(f"⚠️ No items extracted from {f.name}")
        
        if st.session_state.quotes_data and len(st.session_state.quotes_data) > 0:
            if st.button("🗑️ Clear All Quotes"):
                st.session_state.quotes_data = {}
                st.rerun()
    
    st.markdown("---")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🔄 Reset All Data"):
            st.session_state.equipment_schedule = None
            st.session_state.quotes_data = {}
            st.session_state.drawing_filename = None
            st.rerun()
    with col2:
        with st.expander("🔧 Debug Info"):
            st.write(f"Equipment loaded: {st.session_state.equipment_schedule is not None and len(st.session_state.equipment_schedule) > 0 if st.session_state.equipment_schedule else False}")
            st.write(f"Equipment count: {len(st.session_state.equipment_schedule) if st.session_state.equipment_schedule else 0}")
            st.write(f"Quotes loaded: {len(st.session_state.quotes_data) if st.session_state.quotes_data else 0}")

with tabs[1]:
    if st.session_state.equipment_schedule is None or len(st.session_state.equipment_schedule) == 0:
        st.warning("⚠️ Upload drawing first (go to Upload tab)")
    elif st.session_state.quotes_data is None or len(st.session_state.quotes_data) == 0:
        st.warning("⚠️ Upload quotes first (go to Upload tab)")
    else:
        all_q = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_schedule_vs_quotes(st.session_state.equipment_schedule, all_q)
        st.subheader("📊 Coverage Summary")
        act = df[~df['Status'].isin(['IH Supply', 'Existing', 'N/A'])]
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("✓ Quoted", len(act[act['Status'] == '✓ Quoted']))
        c2.metric("❌ Missing", len(act[act['Status'] == '❌ MISSING']))
        c3.metric("⚠ Qty Mismatch", len(act[act['Status'] == '⚠ Qty Mismatch']))
        c4.metric("⚠ Needs Install", len(act[act['Status'] == '⚠ Needs Install']))
        st.metric("💰 Total Quoted", f"${df['Total_Price'].sum():,.2f}")
        ch1, ch2 = st.columns(2)
        with ch1:
            vc = df['Status'].value_counts().reset_index()
            vc.columns = ['Status', 'Count']
            cm = {'✓ Quoted': '#28a745', '❌ MISSING': '#dc3545', '⚠ Qty Mismatch': '#ffc107', '⚠ Needs Install': '#fd7e14', 'IH Supply': '#6c757d', 'Existing': '#adb5bd', 'N/A': '#e9ecef'}
            fig = px.pie(vc, values='Count', names='Status', color='Status', color_discrete_map=cm)
            st.plotly_chart(fig, use_container_width=True)

with tabs[2]:
    if st.session_state.equipment_schedule and len(st.session_state.equipment_schedule) > 0 and st.session_state.quotes_data and len(st.session_state.quotes_data) > 0:
        all_q = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_schedule_vs_quotes(st.session_state.equipment_schedule, all_q)
        st.subheader("📋 Full Report")
        filt = st.multiselect("Filter Status", df['Status'].unique().tolist(), default=df['Status'].unique().tolist())
        fdf = df[df['Status'].isin(filt)]
        def hl(row):
            cm = {'✓ Quoted': 'background-color:#d4edda', '❌ MISSING': 'background-color:#f8d7da', '⚠ Qty Mismatch': 'background-color:#fff3cd', '⚠ Needs Install': 'background-color:#ffe5d0'}
            return [cm.get(row['Status'], '')] * len(row)
        st.dataframe(fdf.style.apply(hl, axis=1), height=500)
        st.subheader("🚨 Critical Missing (Codes 5 & 6)")
        crit = df[(df['Status'] == '❌ MISSING') & (df['Supplier_Code'].isin([5, 6]))]
        if not crit.empty:
            st.dataframe(crit[['No', 'Description', 'Schedule_Qty', 'Supplier_Desc']])
        else:
            st.success("✅ No critical missing!")

with tabs[3]:
    if st.session_state.equipment_schedule and len(st.session_state.equipment_schedule) > 0 and st.session_state.quotes_data and len(st.session_state.quotes_data) > 0:
        all_q = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_schedule_vs_quotes(st.session_state.equipment_schedule, all_q)
        st.subheader("🔢 By Supplier Code")
        summary = []
        for code, desc in SUPPLIER_CODES.items():
            ci = df[df['Supplier_Code'] == code]
            if len(ci) > 0:
                summary.append({'Code': code, 'Description': desc, 'Items': len(ci), 
                               'Quoted': len(ci[ci['Status'].isin(['✓ Quoted'])]),
                               'Missing': len(ci[ci['Status'] == '❌ MISSING']),
                               'Value': ci['Total_Price'].sum()})
        st.dataframe(pd.DataFrame(summary))

with tabs[4]:
    if st.session_state.equipment_schedule and len(st.session_state.equipment_schedule) > 0 and st.session_state.quotes_data and len(st.session_state.quotes_data) > 0:
        all_q = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_schedule_vs_quotes(st.session_state.equipment_schedule, all_q)
        st.subheader("📥 Export")
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as w:
            df.to_excel(w, sheet_name='Analysis', index=False)
            df[df['Status'] == '❌ MISSING'].to_excel(w, sheet_name='Missing', index=False)
            pd.DataFrame(all_q).to_excel(w, sheet_name='All Quotes', index=False)
        out.seek(0)
        st.download_button("📥 Excel Report", out, f"Analysis_{datetime.now().strftime('%Y%m%d')}.xlsx")
    
    if st.session_state.quotes_data and len(st.session_state.quotes_data) > 0:
        st.subheader("📄 Extracted Quotes (Debug)")
        for fn, qs in st.session_state.quotes_data.items():
            with st.expander(f"{fn} - {len(qs)} items"):
                st.dataframe(pd.DataFrame(qs))
    
    st.subheader("🔧 PDF Text Debug")
    dbg = st.file_uploader("Upload PDF", type=['pdf'], key="dbg")
    if dbg and PDF_SUPPORT:
        dbg.seek(0)
        with pdfplumber.open(dbg) as pdf:
            for i, p in enumerate(pdf.pages[:3]):
                with st.expander(f"Page {i+1}"):
                    st.text(p.extract_text())

st.markdown("---")
st.markdown("<center>Equipment Quote Analyzer v5.9 | Drawing No. ↔ Quote Item</center>", unsafe_allow_html=True)
