import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import re
from datetime import datetime

# PDF support - install with: pip install pdfplumber
try:
    import pdfplumber
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False

st.set_page_config(page_title="Drawing Quote Analyzer", page_icon="🔍", layout="wide")

st.markdown("""
<style>
    .main-header { font-size: 2.5rem; font-weight: bold; color: #1f4e79; }
    .sub-header { font-size: 1.2rem; color: #666; }
    div[data-testid="stMetricValue"] { font-size: 1.8rem; }
</style>
""", unsafe_allow_html=True)

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
        st.error("PDF support not available. Install pdfplumber: `pip install pdfplumber`")
        return None, None
    
    text_content = []
    all_tables = []
    
    uploaded_file.seek(0)
    with pdfplumber.open(uploaded_file) as pdf:
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text()
            if text:
                text_content.append(f"--- Page {page_num + 1} ---\n{text}")
            
            tables = page.extract_tables()
            for table in tables:
                if table and len(table) > 1:
                    try:
                        df = pd.DataFrame(table[1:], columns=table[0] if table[0] else None)
                        df['_source_page'] = page_num + 1
                        all_tables.append(df)
                    except:
                        pass
    
    combined_text = "\n".join(text_content)
    return combined_text, all_tables

def parse_excel_file(uploaded_file):
    try:
        uploaded_file.seek(0)
        xl = pd.ExcelFile(uploaded_file)
        all_sheets = {}
        for sheet_name in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name=sheet_name)
            all_sheets[sheet_name] = df
        return all_sheets
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return None

def parse_csv_file(uploaded_file):
    try:
        uploaded_file.seek(0)
        df = pd.read_csv(uploaded_file)
        return {"Sheet1": df}
    except Exception as e:
        st.error(f"Error reading CSV file: {e}")
        return None

def parse_uploaded_file(uploaded_file):
    file_type = uploaded_file.name.split('.')[-1].lower()
    
    if file_type == 'pdf':
        text, tables = parse_pdf_file(uploaded_file)
        return {'type': 'pdf', 'text': text, 'tables': tables, 'filename': uploaded_file.name}
    elif file_type in ['xlsx', 'xls']:
        sheets = parse_excel_file(uploaded_file)
        return {'type': 'excel', 'sheets': sheets, 'filename': uploaded_file.name}
    elif file_type == 'csv':
        sheets = parse_csv_file(uploaded_file)
        return {'type': 'csv', 'sheets': sheets, 'filename': uploaded_file.name}
    else:
        st.error(f"Unsupported file type: {file_type}")
        return None

def extract_equipment_from_pdf_text(text):
    equipment_list = []
    patterns = [
        r'(\d+[a-z]?)\s+([A-Z][A-Z\s\./&\-\(\)]+?)\s+(\d+)\s+(\d+)',
        r'(\d+[a-z]?)\s+([\w\s\./&\-\(\)]+?)\s+(\d+)\s*$',
    ]
    
    lines = text.split('\n')
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        for pattern in patterns:
            match = re.search(pattern, line, re.IGNORECASE)
            if match:
                groups = match.groups()
                item = {
                    'No': groups[0],
                    'Description': groups[1].strip(),
                    'Qty': int(groups[2]) if len(groups) > 2 else 1,
                    'Category': int(groups[3]) if len(groups) > 3 else None
                }
                if len(item['Description']) > 3:
                    equipment_list.append(item)
                break
    
    return equipment_list

def extract_equipment_from_dataframe(df):
    equipment_list = []
    df.columns = df.columns.astype(str).str.strip().str.lower()
    
    col_mapping = {
        'no': ['no', 'no.', 'item', 'item #', 'item#', 'item no', 'item no.', 'number', '#'],
        'description': ['description', 'desc', 'equipment', 'name', 'item description'],
        'qty': ['qty', 'quantity', 'qnty', 'count', 'units'],
        'category': ['category', 'cat', 'supplier code', 'code', 'supplier', 'type']
    }
    
    found_cols = {}
    for key, possibilities in col_mapping.items():
        for col in df.columns:
            col_clean = col.lower().strip()
            if any(p == col_clean for p in possibilities):
                found_cols[key] = col
                break
        if key not in found_cols:
            for col in df.columns:
                if any(p in col.lower() for p in possibilities):
                    found_cols[key] = col
                    break
    
    if 'no' not in found_cols or 'description' not in found_cols:
        return None
    
    for _, row in df.iterrows():
        try:
            no_val = str(row.get(found_cols.get('no', ''), '')).strip()
            desc_val = str(row.get(found_cols.get('description', ''), '')).strip()
            
            if not no_val or not desc_val or no_val.lower() in ['nan', '', 'none']:
                continue
            if desc_val.lower() in ['nan', '', 'none', 'description']:
                continue
            
            qty_val = row.get(found_cols.get('qty', ''), 1)
            try:
                qty_val = int(float(str(qty_val).replace(',', ''))) if pd.notna(qty_val) else 1
            except:
                qty_val = 1
            
            cat_val = row.get(found_cols.get('category', ''), None)
            try:
                cat_val = int(float(str(cat_val))) if pd.notna(cat_val) else None
            except:
                cat_val = None
            
            equipment_list.append({
                'No': no_val,
                'Description': desc_val,
                'Qty': qty_val,
                'Category': cat_val
            })
        except:
            continue
    
    return equipment_list if equipment_list else None

def process_drawing_file(parsed_data):
    equipment_list = []
    
    if parsed_data['type'] == 'pdf':
        if parsed_data['tables']:
            for table_df in parsed_data['tables']:
                extracted = extract_equipment_from_dataframe(table_df)
                if extracted:
                    equipment_list.extend(extracted)
        
        if not equipment_list and parsed_data['text']:
            equipment_list = extract_equipment_from_pdf_text(parsed_data['text'])
    
    elif parsed_data['type'] in ['excel', 'csv']:
        for sheet_name, df in parsed_data['sheets'].items():
            extracted = extract_equipment_from_dataframe(df)
            if extracted:
                equipment_list.extend(extracted)
    
    seen = set()
    unique_list = []
    for item in equipment_list:
        key = (item['No'], item['Description'])
        if key not in seen:
            seen.add(key)
            unique_list.append(item)
    
    return unique_list

def parse_crs_quote_from_text(text, filename):
    """Parse CRS-style quote from PDF text. Quote 'Item' column matches Drawing 'No.' column"""
    quotes = []
    lines = text.split('\n')
    
    current_item = None
    current_qty = 0
    current_desc = ""
    current_unit_price = 0
    current_total_price = 0
    
    main_pattern = re.compile(
        r'^(\d+[a-z]?)\s+(\d+)\s*ea\s+([A-Z][A-Z\s,\./&\-\(\)\']+?)\s+\$?([\d,]+\.?\d*)\s*(?:\$?([\d,]+\.?\d*))?',
        re.IGNORECASE
    )
    
    alt_pattern = re.compile(
        r'^(\d+[a-z]?)\s+(\d+)\s*ea\s+([A-Z][A-Z\s,\./&\-\(\)\']+)',
        re.IGNORECASE
    )
    
    total_pattern = re.compile(r'ITEM\s*TOTAL:\s*\$?([\d,]+\.?\d*)', re.IGNORECASE)
    price_pattern = re.compile(r'\$?([\d,]+\.?\d{2})\s*$')
    nic_pattern = re.compile(r'^(\d+(?:-\d+)?[a-z]?)\s+NIC', re.IGNORECASE)
    
    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue
        
        if nic_pattern.match(line):
            continue
        
        if any(skip in line.lower() for skip in ['page ', 'canadian restaurant', 'bird construc', 'fwg ltc']):
            continue
        
        match = main_pattern.match(line)
        if match:
            if current_item and current_desc:
                quotes.append({
                    'Item': current_item,
                    'Description': current_desc.strip(),
                    'Qty': current_qty,
                    'Unit_Price': current_unit_price,
                    'Total_Price': current_total_price if current_total_price else current_unit_price * current_qty,
                    'Source_File': filename
                })
            
            current_item = match.group(1)
            current_qty = int(match.group(2))
            current_desc = match.group(3).strip()
            
            price1 = match.group(4).replace(',', '')
            current_unit_price = float(price1) if price1 else 0
            
            if match.group(5):
                price2 = match.group(5).replace(',', '')
                current_total_price = float(price2) if price2 else current_unit_price * current_qty
            else:
                current_total_price = current_unit_price * current_qty
            
            continue
        
        alt_match = alt_pattern.match(line)
        if alt_match:
            if current_item and current_desc:
                quotes.append({
                    'Item': current_item,
                    'Description': current_desc.strip(),
                    'Qty': current_qty,
                    'Unit_Price': current_unit_price,
                    'Total_Price': current_total_price if current_total_price else current_unit_price * current_qty,
                    'Source_File': filename
                })
            
            current_item = alt_match.group(1)
            current_qty = int(alt_match.group(2))
            current_desc = alt_match.group(3).strip()
            current_unit_price = 0
            current_total_price = 0
            
            price_match = price_pattern.search(line)
            if price_match:
                current_unit_price = float(price_match.group(1).replace(',', ''))
                current_total_price = current_unit_price * current_qty
            
            continue
        
        total_match = total_pattern.search(line)
        if total_match and current_item:
            current_total_price = float(total_match.group(1).replace(',', ''))
            continue
        
        if current_item and not current_total_price:
            price_match = price_pattern.search(line)
            if price_match and not any(c.isalpha() for c in line.replace('ITEM TOTAL:', '')):
                price_val = float(price_match.group(1).replace(',', ''))
                if not current_unit_price:
                    current_unit_price = price_val
                current_total_price = price_val
    
    if current_item and current_desc:
        quotes.append({
            'Item': current_item,
            'Description': current_desc.strip(),
            'Qty': current_qty,
            'Unit_Price': current_unit_price,
            'Total_Price': current_total_price if current_total_price else current_unit_price * current_qty,
            'Source_File': filename
        })
    
    return quotes

def extract_quotes_from_dataframe(df, filename):
    """Extract quote items from DataFrame. Quote 'Item' matches Drawing 'No.'"""
    quotes = []
    df.columns = df.columns.astype(str).str.strip().str.lower()
    
    col_mapping = {
        'item': ['item', 'item no', 'item no.', 'item #', 'item#', 'no', 'no.', 'line', 'ref', 'number'],
        'description': ['description', 'desc', 'equipment', 'name', 'product'],
        'qty': ['qty', 'quantity', 'qnty', 'count', 'units'],
        'unit_price': ['sell', 'unit price', 'unit', 'price', 'unit cost', 'each'],
        'total_price': ['sell total', 'total', 'total price', 'ext price', 'extended', 'amount', 'ext']
    }
    
    found_cols = {}
    for key, possibilities in col_mapping.items():
        for col in df.columns:
            col_clean = col.lower().strip()
            if any(p == col_clean for p in possibilities):
                found_cols[key] = col
                break
        if key not in found_cols:
            for col in df.columns:
                col_clean = col.lower().strip()
                if any(p in col_clean for p in possibilities):
                    found_cols[key] = col
                    break
    
    for _, row in df.iterrows():
        try:
            desc_val = str(row.get(found_cols.get('description', ''), '')).strip()
            if not desc_val or desc_val.lower() in ['nan', '', 'none', 'description', 'nic']:
                continue
            
            item_val = str(row.get(found_cols.get('item', ''), '')).strip()
            
            qty_val = row.get(found_cols.get('qty', ''), 1)
            qty_str = str(qty_val).lower().replace('ea', '').strip()
            try:
                qty_val = int(float(qty_str.replace(',', ''))) if pd.notna(qty_val) and qty_str else 1
            except:
                qty_val = 1
            
            unit_price = row.get(found_cols.get('unit_price', ''), 0)
            try:
                unit_str = str(unit_price).replace('$', '').replace(',', '').strip()
                unit_price = float(unit_str) if pd.notna(unit_price) and unit_str else 0
            except:
                unit_price = 0
            
            total_price = row.get(found_cols.get('total_price', ''), 0)
            try:
                total_str = str(total_price).replace('$', '').replace(',', '').strip()
                total_price = float(total_str) if pd.notna(total_price) and total_str else 0
            except:
                total_price = unit_price * qty_val if unit_price else 0
            
            quotes.append({
                'Item': item_val,
                'Description': desc_val,
                'Qty': qty_val,
                'Unit_Price': unit_price,
                'Total_Price': total_price,
                'Source_File': filename
            })
        except:
            continue
    
    return quotes

def process_quote_file(parsed_data):
    quotes = []
    
    if parsed_data['type'] == 'pdf':
        if parsed_data['text']:
            quotes = parse_crs_quote_from_text(parsed_data['text'], parsed_data['filename'])
        
        if not quotes and parsed_data['tables']:
            for table_df in parsed_data['tables']:
                extracted = extract_quotes_from_dataframe(table_df, parsed_data['filename'])
                if extracted:
                    quotes.extend(extracted)
    
    elif parsed_data['type'] in ['excel', 'csv']:
        for sheet_name, df in parsed_data['sheets'].items():
            extracted = extract_quotes_from_dataframe(df, parsed_data['filename'])
            if extracted:
                quotes.extend(extracted)
    
    seen_items = set()
    unique_quotes = []
    for q in quotes:
        if q['Item'] not in seen_items:
            seen_items.add(q['Item'])
            unique_quotes.append(q)
    
    return unique_quotes

def match_quote_to_schedule(schedule_item, all_quotes):
    """Match Drawing 'No.' to Quote 'Item'"""
    drawing_no = str(schedule_item['No']).strip().lower()
    desc = schedule_item['Description'].upper()
    
    # PRIMARY: Exact match on Drawing "No." vs Quote "Item"
    for quote in all_quotes:
        quote_item = str(quote.get('Item', '')).strip().lower()
        if drawing_no and quote_item and drawing_no == quote_item:
            return quote
    
    # SECONDARY: Description matching as fallback
    for quote in all_quotes:
        quote_desc = quote['Description'].upper()
        desc_words = set(re.findall(r'\b[A-Z]{3,}\b', desc))
        quote_words = set(re.findall(r'\b[A-Z]{3,}\b', quote_desc))
        common_words = desc_words & quote_words
        
        if len(common_words) >= 2:
            similarity = len(common_words) / max(len(desc_words), len(quote_words), 1)
            if similarity > 0.3:
                return quote
    
    return None

def analyze_schedule_vs_quotes(equipment_schedule, all_quotes):
    analysis = []
    
    for item in equipment_schedule:
        matched_quote = match_quote_to_schedule(item, all_quotes)
        cat = item.get('Category')
        
        if cat in [1, 2, 3]:
            status = "IH Supply"
            issue = "Excluded from analysis - IH handles supply & install"
        elif cat == 8:
            status = "Existing"
            issue = "Existing/relocated equipment"
        elif cat is None:
            status = "N/A"
            issue = "Spare or placeholder item"
        elif matched_quote:
            schedule_qty = item['Qty']
            quote_qty = matched_quote['Qty']
            
            if quote_qty == schedule_qty:
                status = "✓ Quoted"
                issue = None
            elif quote_qty > 0:
                status = "⚠ Qty Mismatch"
                issue = f"Expected {schedule_qty}, got {quote_qty}"
            else:
                status = "⚡ Included"
                issue = "Part of system price"
        else:
            if cat == 7:
                status = "⚠ Needs Install"
                issue = "IH supplies - needs contractor install pricing"
            elif cat in [5, 6]:
                status = "❌ MISSING"
                issue = "Critical - requires contractor quote"
            else:
                status = "❌ MISSING"
                issue = "Not found in any quote"
        
        analysis.append({
            'No': item['No'],
            'Quote_Item': matched_quote['Item'] if matched_quote else '-',
            'Description': item['Description'],
            'Schedule_Qty': item['Qty'],
            'Quote_Qty': matched_quote['Qty'] if matched_quote else 0,
            'Supplier_Code': cat,
            'Supplier_Desc': SUPPLIER_CODES.get(cat, 'Unknown'),
            'Unit_Price': matched_quote['Unit_Price'] if matched_quote else 0,
            'Total_Price': matched_quote['Total_Price'] if matched_quote else 0,
            'Source_File': matched_quote['Source_File'] if matched_quote else '-',
            'Status': status,
            'Issue': issue
        })
    
    return pd.DataFrame(analysis)

# ==================== MAIN UI ====================
st.markdown('<p class="main-header">🔍 Equipment Quote Analyzer</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Upload drawings and quotations (PDF, CSV, XLSX) for analysis</p>', unsafe_allow_html=True)

if not PDF_SUPPORT:
    st.warning("⚠️ **PDF support not installed.** Run: `pip install pdfplumber`")

st.markdown("---")

tabs = st.tabs(["📤 Upload Files", "📊 Analysis Dashboard", "📋 Detailed Report", "🔢 Supplier Summary", "📥 Export"])

with tabs[0]:
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📐 Upload Drawing / Equipment Schedule")
        st.info("Upload equipment schedule. Must have 'No.' column. Formats: PDF, CSV, XLSX")
        
        drawing_file = st.file_uploader("Select Drawing File", type=['pdf', 'csv', 'xlsx', 'xls'], key="drawing_upload")
        
        if drawing_file:
            with st.spinner("Processing drawing file..."):
                parsed = parse_uploaded_file(drawing_file)
                if parsed:
                    equipment = process_drawing_file(parsed)
                    if equipment:
                        st.session_state.equipment_schedule = equipment
                        st.session_state.drawing_filename = drawing_file.name
                        st.success(f"✅ Extracted {len(equipment)} equipment items from {drawing_file.name}")
                        
                        with st.expander("Preview Equipment Schedule", expanded=True):
                            preview_df = pd.DataFrame(equipment)
                            st.dataframe(preview_df, use_container_width=True, height=300)
                    else:
                        st.error("Could not extract equipment schedule. Ensure file has No., Description, Qty columns.")
        
        if st.session_state.equipment_schedule:
            st.markdown(f"**Current Schedule:** {st.session_state.drawing_filename} ({len(st.session_state.equipment_schedule)} items)")
    
    with col2:
        st.subheader("📝 Upload Quotations")
        st.info("Upload quotations. 'Item' column matches Drawing 'No.' column. Formats: PDF, CSV, XLSX")
        
        quote_files = st.file_uploader("Select Quote Files", type=['pdf', 'csv', 'xlsx', 'xls'], key="quote_upload", accept_multiple_files=True)
        
        if quote_files:
            for quote_file in quote_files:
                if quote_file.name not in st.session_state.quotes_data:
                    with st.spinner(f"Processing {quote_file.name}..."):
                        parsed = parse_uploaded_file(quote_file)
                        if parsed:
                            quotes = process_quote_file(parsed)
                            if quotes:
                                st.session_state.quotes_data[quote_file.name] = quotes
                                st.success(f"✅ Extracted {len(quotes)} items from {quote_file.name}")
                                
                                with st.expander(f"Preview: {quote_file.name}", expanded=False):
                                    preview_df = pd.DataFrame(quotes)
                                    st.dataframe(preview_df, use_container_width=True, height=200)
                            else:
                                st.warning(f"⚠️ No quote items found in {quote_file.name}")
        
        if st.session_state.quotes_data:
            st.markdown("**Loaded Quotations:**")
            for fname, quotes in st.session_state.quotes_data.items():
                total_value = sum(q.get('Total_Price', 0) for q in quotes)
                st.markdown(f"- {fname}: {len(quotes)} items (${total_value:,.2f})")
            
            if st.button("🗑️ Clear All Quotes"):
                st.session_state.quotes_data = {}
                st.rerun()

    st.markdown("---")
    if st.button("🔄 Reset All Data"):
        st.session_state.equipment_schedule = None
        st.session_state.quotes_data = {}
        st.session_state.drawing_filename = None
        st.rerun()

with tabs[1]:
    if not st.session_state.equipment_schedule:
        st.warning("⚠️ Please upload a drawing/equipment schedule file first.")
    elif not st.session_state.quotes_data:
        st.warning("⚠️ Please upload at least one quotation file.")
    else:
        all_quotes = []
        for quotes in st.session_state.quotes_data.values():
            all_quotes.extend(quotes)
        
        analysis_df = analyze_schedule_vs_quotes(st.session_state.equipment_schedule, all_quotes)
        
        st.subheader("📊 Quote Coverage Summary")
        actionable = analysis_df[~analysis_df['Status'].isin(['IH Supply', 'Existing', 'N/A'])]
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("✓ Quoted", len(actionable[actionable['Status'] == '✓ Quoted']))
        with col2:
            st.metric("❌ Missing", len(actionable[actionable['Status'] == '❌ MISSING']))
        with col3:
            st.metric("⚠ Qty Mismatch", len(actionable[actionable['Status'] == '⚠ Qty Mismatch']))
        with col4:
            st.metric("⚠ Needs Install", len(actionable[actionable['Status'] == '⚠ Needs Install']))
        
        st.markdown("---")
        col_chart1, col_chart2 = st.columns(2)
        
        with col_chart1:
            status_counts = analysis_df['Status'].value_counts().reset_index()
            status_counts.columns = ['Status', 'Count']
            color_map = {
                '✓ Quoted': '#28a745', '⚡ Included': '#17a2b8', '❌ MISSING': '#dc3545',
                '⚠ Qty Mismatch': '#ffc107', '⚠ Needs Install': '#fd7e14',
                'IH Supply': '#6c757d', 'Existing': '#adb5bd', 'N/A': '#e9ecef'
            }
            fig = px.pie(status_counts, values='Count', names='Status', title='Quote Coverage by Status', color='Status', color_discrete_map=color_map)
            st.plotly_chart(fig, use_container_width=True)
        
        with col_chart2:
            if analysis_df['Supplier_Code'].notna().any():
                supplier_counts = analysis_df.groupby(['Supplier_Code', 'Status']).size().reset_index(name='Count')
                fig2 = px.bar(supplier_counts, x='Supplier_Code', y='Count', color='Status', title='Items by Supplier Code & Status', color_discrete_map=color_map)
                st.plotly_chart(fig2, use_container_width=True)
        
        st.metric("💰 Total Quoted Value", f"${analysis_df['Total_Price'].sum():,.2f}")

with tabs[2]:
    if st.session_state.equipment_schedule and st.session_state.quotes_data:
        all_quotes = []
        for quotes in st.session_state.quotes_data.values():
            all_quotes.extend(quotes)
        
        analysis_df = analyze_schedule_vs_quotes(st.session_state.equipment_schedule, all_quotes)
        
        st.subheader("📋 Full Analysis Report")
        status_filter = st.multiselect("Filter by Status", options=analysis_df['Status'].unique().tolist(), default=analysis_df['Status'].unique().tolist())
        filtered_df = analysis_df[analysis_df['Status'].isin(status_filter)]
        
        def highlight_status(row):
            color_map = {
                '✓ Quoted': 'background-color: #d4edda', '⚡ Included': 'background-color: #d1ecf1',
                '❌ MISSING': 'background-color: #f8d7da', '⚠ Qty Mismatch': 'background-color: #fff3cd',
                '⚠ Needs Install': 'background-color: #ffe5d0'
            }
            return [color_map.get(row['Status'], '')] * len(row)
        
        display_df = filtered_df[['No', 'Quote_Item', 'Description', 'Schedule_Qty', 'Quote_Qty', 'Supplier_Code', 'Unit_Price', 'Total_Price', 'Source_File', 'Status', 'Issue']]
        st.dataframe(display_df.style.apply(highlight_status, axis=1), use_container_width=True, height=500)
        
        st.subheader("🚨 Critical Missing Items (Codes 5 & 6)")
        critical = analysis_df[(analysis_df['Status'] == '❌ MISSING') & (analysis_df['Supplier_Code'].isin([5, 6]))]
        if not critical.empty:
            st.dataframe(critical[['No', 'Description', 'Schedule_Qty', 'Supplier_Desc', 'Issue']], use_container_width=True)
        else:
            st.success("✅ No critical missing items!")
    else:
        st.warning("⚠️ Please upload drawing and quotation files first.")

with tabs[3]:
    if st.session_state.equipment_schedule and st.session_state.quotes_data:
        all_quotes = []
        for quotes in st.session_state.quotes_data.values():
            all_quotes.extend(quotes)
        
        analysis_df = analyze_schedule_vs_quotes(st.session_state.equipment_schedule, all_quotes)
        
        st.subheader("🔢 Summary by Supplier Code")
        summary_data = []
        for code, desc in SUPPLIER_CODES.items():
            code_items = analysis_df[analysis_df['Supplier_Code'] == code]
            if len(code_items) > 0:
                summary_data.append({
                    'Code': code, 'Description': desc, 'Total Items': len(code_items),
                    'Total Qty': code_items['Schedule_Qty'].sum(),
                    'Quoted': len(code_items[code_items['Status'].isin(['✓ Quoted', '⚡ Included'])]),
                    'Missing': len(code_items[code_items['Status'] == '❌ MISSING']),
                    'Qty Issues': len(code_items[code_items['Status'] == '⚠ Qty Mismatch']),
                    'Quoted Value': code_items['Total_Price'].sum()
                })
        
        summary_df = pd.DataFrame(summary_data)
        st.dataframe(summary_df, use_container_width=True)
        
        fig = px.bar(summary_df, x='Code', y=['Quoted', 'Missing', 'Qty Issues'], title='Quote Coverage by Supplier Code', barmode='stack', color_discrete_map={'Quoted': '#28a745', 'Missing': '#dc3545', 'Qty Issues': '#ffc107'})
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("⚠️ Please upload drawing and quotation files first.")

with tabs[4]:
    if st.session_state.equipment_schedule and st.session_state.quotes_data:
        all_quotes = []
        for quotes in st.session_state.quotes_data.values():
            all_quotes.extend(quotes)
        
        analysis_df = analyze_schedule_vs_quotes(st.session_state.equipment_schedule, all_quotes)
        
        st.subheader("📥 Export Reports")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            analysis_df.to_excel(writer, sheet_name='Full Analysis', index=False)
            analysis_df[analysis_df['Status'] == '❌ MISSING'].to_excel(writer, sheet_name='Missing Items', index=False)
            analysis_df[analysis_df['Status'] == '⚠ Qty Mismatch'].to_excel(writer, sheet_name='Qty Mismatch', index=False)
            pd.DataFrame(st.session_state.equipment_schedule).to_excel(writer, sheet_name='Equipment Schedule', index=False)
            pd.DataFrame(all_quotes).to_excel(writer, sheet_name='All Quotes', index=False)
        
        output.seek(0)
        st.download_button("📥 Download Full Excel Report", data=output, file_name=f"Equipment_Quote_Analysis_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.download_button("📥 Download Analysis CSV", data=analysis_df.to_csv(index=False), file_name=f"Quote_Analysis_{datetime.now().strftime('%Y%m%d_%H%M')}.csv", mime="text/csv")
    else:
        st.warning("⚠️ Please upload drawing and quotation files first.")
    
    if st.session_state.quotes_data:
        st.markdown("---")
        st.subheader("📄 Extracted Quote Data (Debug View)")
        for fname, quotes in st.session_state.quotes_data.items():
            with st.expander(f"{fname} - {len(quotes)} items"):
                st.dataframe(pd.DataFrame(quotes), use_container_width=True)

st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; padding: 20px;'>
    <p>Equipment Quote Analyzer v5.3 | Built for Bird Construction</p>
    <p>Matching: Drawing "No." column ↔ Quote "Item" column</p>
</div>
""", unsafe_allow_html=True)
