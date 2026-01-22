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
        for page in pdf.pages:
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
    if parsed_data['type'] == 'pdf':
        if parsed_data.get('text'):
            st.write("**PDF Text Preview (first 3000 chars):**")
            st.text(parsed_data['text'][:3000])
        if parsed_data['tables']:
            st.write(f"**PDF has {len(parsed_data['tables'])} table(s)**")
            for i, tbl in enumerate(parsed_data['tables']):
                st.write(f"Table {i+1} columns: {tbl.columns.tolist()}")
                st.dataframe(tbl.head(5), height=150)
    elif parsed_data['type'] in ['excel', 'csv']:
        if parsed_data.get('sheets'):
            for name, df in parsed_data['sheets'].items():
                st.write(f"**Sheet '{name}'** - Columns: {df.columns.tolist()}")
                st.dataframe(df.head(5), height=150)

def parse_equipment_from_text(text, debug=False):
    """
    Parse equipment schedule from PDF text - handles Zeidler/FWG format
    Columns: No. | NEW EQUIPMENT NUMBER | Description | Qty. | CATEGORY | [electrical specs...]
    """
    equipment_list = []
    lines = text.split('\n')
    
    if debug:
        st.write(f"**Total lines in PDF:** {len(lines)}")
    
    # Skip patterns - headers, notes, etc.
    skip_patterns = [
        r'^EQUIPMENT LIST', r'^CATEGORY', r'^ELECTRICAL', r'^MECHANICAL',
        r'^No\.\s', r'^NEW\s', r'^Description', r'^Load', r'^Volts',
        r'^WATER', r'^WASTE', r'^GAS', r'^EXHAUST', r'^HW', r'^CW',
        r'^E\d+\s', r'^M\d+\s', r'^NOTE:', r'^SUPPLIER CODE',
        r'^\d+\s+IH SUPPLY', r'^\d+\s+CONTRACTOR', r'^\d+\s+EXISTING',
        r'^PROJECT', r'^TITLE', r'^DRAWING', r'^REVISION', r'^zeidler',
        r'^COPYRIGHT', r'^300,', r'^T 403', r'^Zeidler', r'^Interior Health',
        r'^ISSUED', r'^DATE', r'Autodesk Docs', r'^\s*$',
        r'^KITCHEN EQUIPMENT', r'^1 : \d+', r'^K-\d+',
        r'^THIS PLAN', r'^ALL SERVICES', r'^ELECTRICAL CONTRACTOR',
        r'^MECHANICAL CONTRACTOR', r'^KITCHEN CONTRACTOR',
        r'^UPON COMPLETION', r'^AT THIS POINT', r'^PROJECT ADDRESS',
        r'^Elec\. RI', r'^Height', r'^Direct', r'^Indirect',
        r'^Conn\. Type', r'^Ph\.', r'^Load\s+MBH'
    ]
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        # Skip header/note lines
        skip = False
        for pattern in skip_patterns:
            if re.match(pattern, line, re.IGNORECASE):
                skip = True
                break
        if skip:
            continue
        
        # Equipment line pattern: starts with item number (1, 1a, 2, 3, 9a, etc.)
        item_match = re.match(r'^(\d+[a-z]?)\s+(.+)$', line, re.IGNORECASE)
        if not item_match:
            continue
        
        item_no = item_match.group(1)
        rest = item_match.group(2).strip()
        
        # Skip supplier code definitions (e.g., "1 IH SUPPLY / IH INSTALL")
        if re.match(r'^IH SUPPLY|^CONTRACTOR|^EXISTING', rest, re.IGNORECASE):
            continue
        
        # Check for NEW EQUIPMENT NUMBER at start (e.g., 9038, 1195.12, 1302.15)
        equip_num = None
        equip_match = re.match(r'^(\d+\.?\d*)\s+(.+)$', rest)
        if equip_match:
            potential_equip = equip_match.group(1)
            # Equipment numbers are typically 4+ digits or have decimal (like 1195.12)
            if len(potential_equip) >= 4 or '.' in potential_equip:
                equip_num = potential_equip
                rest = equip_match.group(2).strip()
        
        # Also check for "-" as equipment number (meaning none)
        if rest.startswith('- '):
            equip_num = '-'
            rest = rest[2:].strip()
        
        # Now parse: DESCRIPTION QTY CATEGORY [ELECTRICAL_SPECS...]
        # Electrical specs start with patterns like: 12A, 0.3KW, JUNCTION, RECEPTACLE, etc.
        
        # Find where electrical specs begin
        elec_patterns = [
            r'\d+\.?\d*A\s+\d+V',      # e.g., "12A 120V", "14A 208V"
            r'\d+\.?\d*KW\s+\d+V',     # e.g., "0.3KW 120V"
            r'\d+\.?\d*A\s+\d+\.?\d*KW', # variations
            r'\bJUNCTION\b',
            r'\bRECEPTACLE\b',
            r'\bSEE NOTE\b',
            r'\bTWO SERVICES\b',
            r'\bSERVICES REQ',
            r'\bLIGHTS[;,]',
            r'\d+\s+FFD\b',             # Drain specs
            r'\bSTUB-UP\b',
            r'\bWASTE TO\b',
            r'\bX\s+X\s+X',             # placeholder specs
        ]
        
        elec_start = len(rest)
        for pattern in elec_patterns:
            m = re.search(pattern, rest, re.IGNORECASE)
            if m and m.start() < elec_start:
                elec_start = m.start()
        
        before_elec = rest[:elec_start].strip()
        
        # Parse: DESCRIPTION QTY CATEGORY from before_elec
        # Category is 1-8 or "-", Qty is typically 1-9 (could be larger)
        # Pattern: "DESCRIPTION QTY CATEGORY" where both are at the end
        
        # Try to match qty and category at the end
        # Format: "... DESCRIPTION X Y" where X=qty (1-99), Y=category (1-8 or -)
        qty_cat_match = re.search(r'\s+(\d+)\s+([1-8]|-)\s*$', before_elec)
        
        if qty_cat_match:
            description = before_elec[:qty_cat_match.start()].strip()
            qty = int(qty_cat_match.group(1))
            cat_str = qty_cat_match.group(2)
            category = int(cat_str) if cat_str != '-' else None
        else:
            # Try with just qty (category might be missing or "-")
            qty_match = re.search(r'\s+(\d+)\s*$', before_elec)
            if qty_match:
                description = before_elec[:qty_match.start()].strip()
                qty = int(qty_match.group(1))
                category = None
            else:
                if debug:
                    st.write(f"⚠️ Could not parse qty/cat: {line[:80]}...")
                continue
        
        # Clean up description
        description = re.sub(r'\s+', ' ', description).strip()
        
        # Skip invalid entries
        if len(description) < 2:
            continue
        
        # Handle SPARE items
        if description.upper() == 'SPARE' or description == '-':
            description = 'SPARE'
        
        equipment_list.append({
            'No': item_no,
            'Equip_Num': equip_num if equip_num else '-',
            'Description': description,
            'Qty': qty,
            'Category': category
        })
        
        if debug and len(equipment_list) <= 10:
            st.write(f"✅ No={item_no}, Equip={equip_num}, Desc={description[:35]}..., Qty={qty}, Cat={category}")
    
    if debug:
        st.write(f"**Total equipment items extracted:** {len(equipment_list)}")
    
    return equipment_list

def extract_equipment_from_dataframe(df, debug=False):
    """Extract equipment items from a dataframe"""
    equipment_list = []
    original_cols = df.columns.tolist()
    df = df.dropna(how='all').reset_index(drop=True)
    df.columns = df.columns.astype(str).str.strip().str.lower()
    
    if debug:
        st.write("**Original columns:**", original_cols)
        st.write("**Normalized columns:**", df.columns.tolist())
    
    col_map = {
        'no': ['no', 'no.', 'item', 'item #', 'item no', 'item no.', 'number', '#'],
        'equip_num': ['new equipment number', 'equipment number', 'equip num', 'equip no', 'equip no.', 'equipment no', 'equipment no.'],
        'description': ['description', 'desc', 'equipment', 'name', 'item description'],
        'qty': ['qty', 'qty.', 'quantity', 'count'],
        'category': ['category', 'cat', 'supplier code', 'code', 'supplier']
    }
    
    found = {}
    for key, opts in col_map.items():
        for col in df.columns:
            col_clean = col.lower().strip()
            if col_clean in opts:
                found[key] = col
                break
        if key not in found:
            for col in df.columns:
                col_clean = col.lower().strip()
                if any(o in col_clean for o in opts):
                    found[key] = col
                    break
    
    if debug:
        st.write("**Found column mapping:**", found)
    
    # Fallback for no/description if not found
    if 'no' not in found or 'description' not in found:
        if len(df.columns) >= 3:
            sample = df[df.columns[0]].dropna().head(10).astype(str).tolist()
            if any(re.match(r'^\d+[a-zA-Z]?$', str(v).strip()) for v in sample if str(v).strip()):
                if 'no' not in found:
                    found['no'] = df.columns[0]
                if 'equip_num' not in found:
                    found['equip_num'] = df.columns[1]
                if 'description' not in found:
                    found['description'] = df.columns[2]
    
    if 'no' not in found or 'description' not in found:
        if debug:
            st.error(f"Missing required columns. Found: {found}")
        return None
    
    for idx, row in df.iterrows():
        try:
            no = str(row.get(found.get('no', ''), '')).strip()
            desc = str(row.get(found.get('description', ''), '')).strip()
            
            if not no or no.lower() in ['nan', '', 'none', 'no', 'no.', 'item']:
                continue
            if not desc or desc.lower() in ['nan', '', 'none', 'description']:
                continue
            
            equip_num = '-'
            if 'equip_num' in found:
                en = str(row.get(found['equip_num'], '')).strip()
                equip_num = en if en and en.lower() not in ['nan', '', 'none'] else '-'
            
            qty = 1
            if 'qty' in found:
                try:
                    qval = str(row.get(found['qty'], 1)).replace(',', '').strip()
                    qval = re.sub(r'[^\d.]', '', qval)
                    qty = int(float(qval)) if qval else 1
                except:
                    pass
            
            cat = None
            if 'category' in found:
                try:
                    cval = str(row.get(found['category'], '')).strip()
                    cval = re.sub(r'[^\d]', '', cval)
                    cat = int(float(cval)) if cval else None
                except:
                    pass
            
            equipment_list.append({
                'No': no, 
                'Equip_Num': equip_num,
                'Description': desc, 
                'Qty': qty, 
                'Category': cat
            })
        except Exception as e:
            if debug:
                st.write(f"Row {idx} error: {e}")
            continue
    
    if debug:
        st.write(f"**Extracted {len(equipment_list)} items from dataframe**")
    
    return equipment_list if equipment_list else None

def process_drawing_file(parsed_data, debug=False):
    """Process drawing file and extract equipment list"""
    equipment_list = []
    
    if parsed_data['type'] == 'pdf':
        # Try text parsing first for PDFs
        if parsed_data.get('text'):
            if debug:
                st.write("**Attempting text-based extraction...**")
            text_equip = parse_equipment_from_text(parsed_data['text'], debug=debug)
            if text_equip and len(text_equip) > 0:
                equipment_list.extend(text_equip)
                if debug:
                    st.success(f"Text extraction found {len(text_equip)} items")
        
        # If text parsing got few results, try table extraction
        if len(equipment_list) < 5 and parsed_data.get('tables'):
            if debug:
                st.write(f"**Trying table extraction ({len(parsed_data['tables'])} tables)...**")
            for i, tbl in enumerate(parsed_data['tables']):
                if debug:
                    st.write(f"Processing Table {i+1}")
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
    """Match quote to schedule item by No"""
    no = str(item['No']).strip().lower()
    for q in quotes:
        qi = str(q.get('Item', '')).strip().lower()
        if no == qi:
            return q
    # Try numeric match (ignore letter suffix)
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
    """Analyze equipment schedule against quotes"""
    analysis = []
    for item in schedule:
        match = match_quote_to_schedule(item, quotes)
        cat = item.get('Category')
        
        # Determine status based on category and quote match
        if cat in [1, 2, 3]:
            status, issue = "IH Supply", "IH handles supply & install"
        elif cat == 8:
            status, issue = "Existing", "Existing/relocated"
        elif cat is None or item.get('Description', '').upper() == 'SPARE':
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
            'Equip_Num': item.get('Equip_Num', '-'),
            'Quote_Item': match['Item'] if match else '-',
            'Description': item['Description'],
            'Schedule_Qty': item['Qty'],
            'Quote_Qty': match['Qty'] if match else 0,
            'Supplier_Code': cat,
            'Supplier_Desc': SUPPLIER_CODES.get(cat, 'Unknown') if cat else 'N/A',
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
        debug_mode = st.checkbox("🔧 Debug Mode (show parsing details)", key="debug_draw")
        
        if st.session_state.equipment_schedule and len(st.session_state.equipment_schedule) > 0:
            st.success(f"✅ Loaded: {st.session_state.drawing_filename} ({len(st.session_state.equipment_schedule)} items)")
            with st.expander("View Equipment Schedule"):
                df_preview = pd.DataFrame(st.session_state.equipment_schedule)
                st.dataframe(df_preview, height=300)
        
        df_file = st.file_uploader("Select Drawing File", type=['pdf', 'csv', 'xlsx', 'xls'], key="draw")
        if df_file:
            with st.spinner("Processing drawing..."):
                parsed = parse_uploaded_file(df_file)
                if parsed:
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
**Tips:**
1. Enable **Debug Mode** to see parsing details
2. For PDFs: Equipment list should have format: No. | Equip# | Description | Qty | Category
3. For Excel: Need columns like `No.`, `Description`, `Qty.`, `CATEGORY`
                        """)
                        if not debug_mode:
                            with st.expander("📋 Click to see file structure"):
                                show_file_preview(parsed)
                else:
                    st.error("❌ Could not parse file.")
    
    with c2:
        st.subheader("📝 Quotations")
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
        with st.expander("🔧 Session Info"):
            st.write(f"Equipment count: {len(st.session_state.equipment_schedule) if st.session_state.equipment_schedule else 0}")
            st.write(f"Quote files: {len(st.session_state.quotes_data) if st.session_state.quotes_data else 0}")

with tabs[1]:
    if not st.session_state.equipment_schedule or len(st.session_state.equipment_schedule) == 0:
        st.warning("⚠️ Upload drawing first (go to Upload tab)")
    elif not st.session_state.quotes_data or len(st.session_state.quotes_data) == 0:
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
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("💰 Total Quoted Value", f"${df['Total_Price'].sum():,.2f}")
        with col2:
            total_items = len(df)
            actionable = len(act)
            st.metric("📦 Total Items", f"{total_items} ({actionable} actionable)")
        
        ch1, ch2 = st.columns(2)
        with ch1:
            vc = df['Status'].value_counts().reset_index()
            vc.columns = ['Status', 'Count']
            cm = {'✓ Quoted': '#28a745', '❌ MISSING': '#dc3545', '⚠ Qty Mismatch': '#ffc107', 
                  '⚠ Needs Install': '#fd7e14', 'IH Supply': '#6c757d', 'Existing': '#adb5bd', 'N/A': '#e9ecef'}
            fig = px.pie(vc, values='Count', names='Status', color='Status', color_discrete_map=cm, title="Status Distribution")
            st.plotly_chart(fig, use_container_width=True)
        
        with ch2:
            # Category breakdown
            cat_summary = df.groupby('Supplier_Code').agg({'No': 'count', 'Total_Price': 'sum'}).reset_index()
            cat_summary.columns = ['Code', 'Items', 'Value']
            cat_summary = cat_summary[cat_summary['Code'].notna()]
            cat_summary['Code'] = cat_summary['Code'].astype(int)
            cat_summary['Label'] = cat_summary['Code'].map(lambda x: f"Code {x}")
            fig2 = px.bar(cat_summary, x='Label', y='Items', title="Items by Supplier Code", color='Value')
            st.plotly_chart(fig2, use_container_width=True)

with tabs[2]:
    if st.session_state.equipment_schedule and len(st.session_state.equipment_schedule) > 0 and st.session_state.quotes_data and len(st.session_state.quotes_data) > 0:
        all_q = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_schedule_vs_quotes(st.session_state.equipment_schedule, all_q)
        
        st.subheader("📋 Full Analysis Report")
        
        # Filters
        col1, col2 = st.columns(2)
        with col1:
            filt_status = st.multiselect("Filter by Status", df['Status'].unique().tolist(), default=df['Status'].unique().tolist())
        with col2:
            filt_cat = st.multiselect("Filter by Supplier Code", sorted([c for c in df['Supplier_Code'].unique() if c is not None]), 
                                      default=sorted([c for c in df['Supplier_Code'].unique() if c is not None]))
        
        fdf = df[df['Status'].isin(filt_status)]
        if filt_cat:
            fdf = fdf[(fdf['Supplier_Code'].isin(filt_cat)) | (fdf['Supplier_Code'].isna())]
        
        def highlight_status(row):
            cm = {'✓ Quoted': 'background-color:#d4edda', '❌ MISSING': 'background-color:#f8d7da', 
                  '⚠ Qty Mismatch': 'background-color:#fff3cd', '⚠ Needs Install': 'background-color:#ffe5d0'}
            return [cm.get(row['Status'], '')] * len(row)
        
        st.dataframe(fdf.style.apply(highlight_status, axis=1), height=500, use_container_width=True)
        
        # Critical Missing section
        st.subheader("🚨 Critical Missing (Supplier Codes 5 & 6)")
        crit = df[(df['Status'] == '❌ MISSING') & (df['Supplier_Code'].isin([5, 6]))]
        if not crit.empty:
            st.error(f"Found {len(crit)} critical missing items that require contractor quotes!")
            st.dataframe(crit[['No', 'Equip_Num', 'Description', 'Schedule_Qty', 'Supplier_Code', 'Supplier_Desc']], use_container_width=True)
        else:
            st.success("✅ No critical missing items!")

with tabs[3]:
    if st.session_state.equipment_schedule and len(st.session_state.equipment_schedule) > 0 and st.session_state.quotes_data and len(st.session_state.quotes_data) > 0:
        all_q = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_schedule_vs_quotes(st.session_state.equipment_schedule, all_q)
        
        st.subheader("🔢 Summary by Supplier Code")
        summary = []
        for code, desc in SUPPLIER_CODES.items():
            ci = df[df['Supplier_Code'] == code]
            if len(ci) > 0:
                summary.append({
                    'Code': code, 
                    'Description': desc, 
                    'Total Items': len(ci),
                    'Quoted': len(ci[ci['Status'] == '✓ Quoted']),
                    'Missing': len(ci[ci['Status'] == '❌ MISSING']),
                    'Needs Install': len(ci[ci['Status'] == '⚠ Needs Install']),
                    'Total Value': f"${ci['Total_Price'].sum():,.2f}"
                })
        
        st.dataframe(pd.DataFrame(summary), use_container_width=True)
        
        # Items with no category
        no_cat = df[df['Supplier_Code'].isna()]
        if len(no_cat) > 0:
            st.subheader("📋 Items with No Category (SPARE/Placeholder)")
            st.dataframe(no_cat[['No', 'Equip_Num', 'Description', 'Schedule_Qty']], use_container_width=True)

with tabs[4]:
    st.subheader("📥 Export Data")
    
    if st.session_state.equipment_schedule and len(st.session_state.equipment_schedule) > 0:
        # Export equipment schedule only
        st.write("**Equipment Schedule:**")
        eq_df = pd.DataFrame(st.session_state.equipment_schedule)
        out_eq = io.BytesIO()
        eq_df.to_excel(out_eq, index=False)
        out_eq.seek(0)
        st.download_button("📥 Equipment Schedule (Excel)", out_eq, f"Equipment_{datetime.now().strftime('%Y%m%d')}.xlsx")
    
    if st.session_state.equipment_schedule and len(st.session_state.equipment_schedule) > 0 and st.session_state.quotes_data and len(st.session_state.quotes_data) > 0:
        all_q = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_schedule_vs_quotes(st.session_state.equipment_schedule, all_q)
        
        st.write("**Full Analysis Report:**")
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as w:
            df.to_excel(w, sheet_name='Full Analysis', index=False)
            df[df['Status'] == '❌ MISSING'].to_excel(w, sheet_name='Missing Items', index=False)
            df[(df['Status'] == '❌ MISSING') & (df['Supplier_Code'].isin([5, 6]))].to_excel(w, sheet_name='Critical Missing', index=False)
            pd.DataFrame(all_q).to_excel(w, sheet_name='All Quotes', index=False)
        out.seek(0)
        st.download_button("📥 Full Analysis Report (Excel)", out, f"Analysis_{datetime.now().strftime('%Y%m%d')}.xlsx")
    
    # Debug section
    if st.session_state.quotes_data and len(st.session_state.quotes_data) > 0:
        st.subheader("📄 Extracted Quotes (Debug)")
        for fn, qs in st.session_state.quotes_data.items():
            with st.expander(f"{fn} - {len(qs)} items"):
                st.dataframe(pd.DataFrame(qs))
    
    st.subheader("🔧 PDF Text Debug")
    dbg = st.file_uploader("Upload PDF to view raw text", type=['pdf'], key="dbg")
    if dbg and PDF_SUPPORT:
        dbg.seek(0)
        with pdfplumber.open(dbg) as pdf:
            for i, p in enumerate(pdf.pages[:3]):
                with st.expander(f"Page {i+1}"):
                    st.text(p.extract_text())

st.markdown("---")
st.markdown("<center>Equipment Quote Analyzer v6.1 | Zeidler/FWG Format Support</center>", unsafe_allow_html=True)
