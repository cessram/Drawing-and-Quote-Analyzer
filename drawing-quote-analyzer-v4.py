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

# Pre-compile regex patterns for performance
ITEM_PATTERN = re.compile(r'^(\d+[a-z]?)\s+(.+)$', re.IGNORECASE)
EQUIP_NUM_PATTERN = re.compile(r'^(\d+\.?\d*)\s+(.+)$')
QTY_CAT_PATTERN = re.compile(r'\s+(\d+)\s+([1-8]|-)\s*$')
QTY_ONLY_PATTERN = re.compile(r'\s+(\d+)\s*$')
SKIP_SUPPLIER = re.compile(r'^(IH SUPPLY|CONTRACTOR|EXISTING)', re.IGNORECASE)
ELEC_PATTERN = re.compile(r'(\d+\.?\d*A\s+\d+V|\d+\.?\d*KW|JUNCTION|RECEPTACLE|SEE NOTE|TWO SERVICES|SERVICES REQ|LIGHTS[;,]|\d+\s+FFD|STUB-UP|WASTE TO|X\s+X\s+X)', re.IGNORECASE)

SKIP_WORDS = frozenset(['equipment list', 'category', 'electrical', 'mechanical', 'description', 
    'load', 'volts', 'water', 'waste', 'gas', 'exhaust', 'project', 'title', 'drawing', 
    'revision', 'zeidler', 'copyright', 'issued', 'date', 'kitchen equipment', 'this plan',
    'all services', 'electrical contractor', 'mechanical contractor', 'kitchen contractor',
    'upon completion', 'at this point', 'project address', 'supplier code', 'note:',
    'interior health', 'autodesk docs'])

if 'equipment_schedule' not in st.session_state:
    st.session_state.equipment_schedule = None
if 'quotes_data' not in st.session_state:
    st.session_state.quotes_data = {}
if 'drawing_filename' not in st.session_state:
    st.session_state.drawing_filename = None

def parse_pdf_file(uploaded_file):
    if not PDF_SUPPORT:
        return None, None
    text_content = []
    uploaded_file.seek(0)
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                text_content.append(text)
    return "\n".join(text_content), None

def parse_excel_file(uploaded_file):
    try:
        uploaded_file.seek(0)
        xl = pd.ExcelFile(uploaded_file)
        return {name: pd.read_excel(xl, sheet_name=name) for name in xl.sheet_names}
    except:
        return None

def parse_csv_file(uploaded_file):
    try:
        uploaded_file.seek(0)
        return {"Sheet1": pd.read_csv(uploaded_file)}
    except:
        return None

def parse_uploaded_file(uploaded_file):
    ext = uploaded_file.name.split('.')[-1].lower()
    if ext == 'pdf':
        text, _ = parse_pdf_file(uploaded_file)
        return {'type': 'pdf', 'text': text, 'filename': uploaded_file.name}
    elif ext in ['xlsx', 'xls']:
        sheets = parse_excel_file(uploaded_file)
        return {'type': 'excel', 'sheets': sheets, 'filename': uploaded_file.name}
    elif ext == 'csv':
        sheets = parse_csv_file(uploaded_file)
        return {'type': 'csv', 'sheets': sheets, 'filename': uploaded_file.name}
    return None

def parse_equipment_from_text(text):
    """Parse equipment schedule from PDF text - optimized for Zeidler/FWG format"""
    equipment_list = []
    lines = text.split('\n')
    
    for line in lines:
        line = line.strip()
        if not line or len(line) < 5:
            continue
        
        # Quick skip check
        line_lower = line.lower()
        if any(sw in line_lower for sw in SKIP_WORDS):
            continue
        if line_lower.startswith(('e1 ', 'e2 ', 'e3 ', 'e4 ', 'e5 ', 'e6 ', 'e7 ', 'm1 ', 'm2 ', 'm3 ', 'm4 ', 'm5 ', 'm6 ')):
            continue
        if line.startswith(('No.', 'NEW ', '1 :', 'K-', '300,', 'T 403')):
            continue
        
        # Match equipment line
        item_match = ITEM_PATTERN.match(line)
        if not item_match:
            continue
        
        item_no = item_match.group(1)
        rest = item_match.group(2).strip()
        
        # Skip supplier code definitions
        if SKIP_SUPPLIER.match(rest):
            continue
        
        # Check for equipment number
        equip_num = '-'
        equip_match = EQUIP_NUM_PATTERN.match(rest)
        if equip_match:
            potential = equip_match.group(1)
            if len(potential) >= 4 or '.' in potential:
                equip_num = potential
                rest = equip_match.group(2).strip()
        
        if rest.startswith('- '):
            equip_num = '-'
            rest = rest[2:].strip()
        
        # Find electrical specs start
        elec_match = ELEC_PATTERN.search(rest)
        before_elec = rest[:elec_match.start()].strip() if elec_match else rest
        
        # Parse qty and category
        qty_cat = QTY_CAT_PATTERN.search(before_elec)
        if qty_cat:
            description = before_elec[:qty_cat.start()].strip()
            qty = int(qty_cat.group(1))
            cat_str = qty_cat.group(2)
            category = int(cat_str) if cat_str != '-' else None
        else:
            qty_only = QTY_ONLY_PATTERN.search(before_elec)
            if qty_only:
                description = before_elec[:qty_only.start()].strip()
                qty = int(qty_only.group(1))
                category = None
            else:
                continue
        
        description = ' '.join(description.split())
        if len(description) < 2:
            continue
        if description.upper() in ('SPARE', '-'):
            description = 'SPARE'
        
        equipment_list.append({
            'No': item_no,
            'Equip_Num': equip_num,
            'Description': description,
            'Qty': qty,
            'Category': category
        })
    
    return equipment_list

def extract_equipment_from_dataframe(df):
    """Extract equipment from Excel/CSV"""
    equipment_list = []
    df = df.dropna(how='all').reset_index(drop=True)
    df.columns = df.columns.astype(str).str.strip().str.lower()
    
    col_map = {
        'no': ['no', 'no.', 'item', 'item #', 'item no', 'item no.', 'number', '#'],
        'equip_num': ['new equipment number', 'equipment number', 'equip num', 'equip no'],
        'description': ['description', 'desc', 'equipment', 'name', 'item description'],
        'qty': ['qty', 'qty.', 'quantity', 'count'],
        'category': ['category', 'cat', 'supplier code', 'code', 'supplier']
    }
    
    found = {}
    for key, opts in col_map.items():
        for col in df.columns:
            if col.lower().strip() in opts or any(o in col.lower() for o in opts):
                found[key] = col
                break
    
    if 'no' not in found or 'description' not in found:
        if len(df.columns) >= 3:
            found['no'] = df.columns[0]
            found['equip_num'] = df.columns[1]
            found['description'] = df.columns[2]
    
    if 'no' not in found or 'description' not in found:
        return None
    
    for _, row in df.iterrows():
        try:
            no = str(row.get(found['no'], '')).strip()
            desc = str(row.get(found['description'], '')).strip()
            if not no or no.lower() in ('nan', '', 'no', 'no.', 'item') or not desc or desc.lower() in ('nan', '', 'description'):
                continue
            
            equip_num = str(row.get(found.get('equip_num', ''), '-')).strip()
            equip_num = equip_num if equip_num and equip_num.lower() not in ('nan', '') else '-'
            
            qty = 1
            if 'qty' in found:
                try:
                    qty = int(float(re.sub(r'[^\d.]', '', str(row.get(found['qty'], 1)))))
                except:
                    pass
            
            cat = None
            if 'category' in found:
                try:
                    cat = int(float(re.sub(r'[^\d]', '', str(row.get(found['category'], '')))))
                except:
                    pass
            
            equipment_list.append({'No': no, 'Equip_Num': equip_num, 'Description': desc, 'Qty': qty, 'Category': cat})
        except:
            continue
    
    return equipment_list if equipment_list else None

def process_drawing_file(parsed_data):
    """Process drawing file and extract equipment list"""
    equipment_list = []
    
    if parsed_data['type'] == 'pdf' and parsed_data.get('text'):
        equipment_list = parse_equipment_from_text(parsed_data['text'])
    elif parsed_data['type'] in ['excel', 'csv'] and parsed_data.get('sheets'):
        for df in parsed_data['sheets'].values():
            ext = extract_equipment_from_dataframe(df)
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
    current_item = current_desc = None
    current_qty = current_unit = current_total = 0
    
    skip_words = ('page ', 'canadian restaurant', 'bird construc', 'fwg ltc', 
                  'item qty description', 'sell total', 'merchandise', 'prices are in', 'quote valid')
    
    for line in text.split('\n'):
        line = line.strip()
        if not line:
            continue
        ll = line.lower()
        if any(s in ll for s in skip_words) or re.match(r'^[\d\-]+[a-z]?\s+NIC', line, re.IGNORECASE):
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
                quotes.append({'Item': current_item, 'Description': current_desc.strip(), 'Qty': current_qty,
                    'Unit_Price': current_unit, 'Total_Price': current_total or current_unit * current_qty, 'Source_File': filename})
            current_item, current_qty = m.group(1), int(m.group(2))
            rest = m.group(3).strip()
            prices = re.findall(r'\$?([\d,]+\.\d{2})', rest)
            current_desc = re.sub(r'\s*\$?[\d,]+\.\d{2}', '', rest).strip()
            current_unit = float(prices[0].replace(',', '')) if prices else 0
            current_total = float(prices[1].replace(',', '')) if len(prices) >= 2 else current_unit * current_qty
    
    if current_item and current_desc:
        quotes.append({'Item': current_item, 'Description': current_desc.strip(), 'Qty': current_qty,
            'Unit_Price': current_unit, 'Total_Price': current_total or current_unit * current_qty, 'Source_File': filename})
    return quotes

def extract_quotes_from_dataframe(df, filename):
    quotes = []
    df.columns = df.columns.astype(str).str.strip().str.lower()
    col_map = {'item': ['item', 'item no', 'no', 'no.'], 'description': ['description', 'desc', 'equipment', 'name'],
        'qty': ['qty', 'quantity'], 'unit_price': ['sell', 'unit price', 'price'], 'total_price': ['sell total', 'total', 'total price', 'amount']}
    found = {k: next((c for c in df.columns if any(o == c.lower().strip() for o in v)), None) for k, v in col_map.items()}
    
    for _, row in df.iterrows():
        try:
            desc = str(row.get(found.get('description', ''), '')).strip()
            if not desc or desc.lower() in ('nan', '', 'nic'):
                continue
            item = str(row.get(found.get('item', ''), '')).strip()
            qty = int(float(str(row.get(found.get('qty', ''), 1)).replace('ea', '').replace(',', '').strip() or 1))
            up = float(str(row.get(found.get('unit_price', ''), 0)).replace('$', '').replace(',', '').strip() or 0)
            tp = float(str(row.get(found.get('total_price', ''), 0)).replace('$', '').replace(',', '').strip() or 0) or up * qty
            quotes.append({'Item': item, 'Description': desc, 'Qty': qty, 'Unit_Price': up, 'Total_Price': tp, 'Source_File': filename})
        except:
            continue
    return quotes

def process_quote_file(parsed_data):
    quotes = []
    if parsed_data['type'] == 'pdf' and parsed_data.get('text'):
        quotes = parse_crs_quote_from_text(parsed_data['text'], parsed_data['filename'])
    elif parsed_data['type'] in ['excel', 'csv'] and parsed_data.get('sheets'):
        for df in parsed_data['sheets'].values():
            quotes.extend(extract_quotes_from_dataframe(df, parsed_data['filename']))
    
    seen = set()
    return [q for q in quotes if q['Item'] and q['Item'] not in seen and not seen.add(q['Item'])]

def match_quote_to_schedule(item, quotes):
    no = str(item['No']).strip().lower()
    for q in quotes:
        if str(q.get('Item', '')).strip().lower() == no:
            return q
    try:
        no_int = int(re.sub(r'[a-zA-Z]', '', no))
        for q in quotes:
            try:
                if int(re.sub(r'[a-zA-Z]', '', str(q.get('Item', '')).strip())) == no_int:
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
        elif cat is None or item.get('Description', '').upper() == 'SPARE':
            status, issue = "N/A", "Spare or placeholder"
        elif match:
            status, issue = ("✓ Quoted", None) if match['Qty'] == item['Qty'] else ("⚠ Qty Mismatch", f"Expected {item['Qty']}, got {match['Qty']}")
        else:
            if cat == 7:
                status, issue = "⚠ Needs Install", "IH supplies - needs install pricing"
            elif cat in [5, 6]:
                status, issue = "❌ MISSING", "Critical - requires quote"
            else:
                status, issue = "❌ MISSING", "Not found"
        
        analysis.append({
            'No': item['No'], 'Equip_Num': item.get('Equip_Num', '-'), 'Quote_Item': match['Item'] if match else '-',
            'Description': item['Description'], 'Schedule_Qty': item['Qty'], 'Quote_Qty': match['Qty'] if match else 0,
            'Supplier_Code': cat, 'Supplier_Desc': SUPPLIER_CODES.get(cat, 'N/A') if cat else 'N/A',
            'Unit_Price': match['Unit_Price'] if match else 0, 'Total_Price': match['Total_Price'] if match else 0,
            'Source_File': match['Source_File'] if match else '-', 'Status': status, 'Issue': issue
        })
    return pd.DataFrame(analysis)

# ==================== UI ====================
st.markdown("## 🔍 Equipment Quote Analyzer")

if not PDF_SUPPORT:
    st.warning("PDF support not installed. Run: pip install pdfplumber")

tabs = st.tabs(["📤 Upload", "📊 Dashboard", "📋 Report", "🔢 Summary", "📥 Export"])

with tabs[0]:
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("📐 Drawing / Equipment Schedule")
        if st.session_state.equipment_schedule:
            st.success(f"✅ {st.session_state.drawing_filename} ({len(st.session_state.equipment_schedule)} items)")
            with st.expander("View Equipment"):
                st.dataframe(pd.DataFrame(st.session_state.equipment_schedule), height=250)
        
        df_file = st.file_uploader("Select Drawing", type=['pdf', 'csv', 'xlsx', 'xls'], key="draw")
        if df_file and df_file.name != st.session_state.drawing_filename:
            with st.spinner("Processing..."):
                parsed = parse_uploaded_file(df_file)
                if parsed:
                    equip = process_drawing_file(parsed)
                    if equip:
                        st.session_state.equipment_schedule = equip
                        st.session_state.drawing_filename = df_file.name
                        st.rerun()
                    else:
                        st.error("❌ Could not extract equipment. Check file format.")
    
    with c2:
        st.subheader("📝 Quotations")
        if st.session_state.quotes_data:
            st.success(f"✅ {len(st.session_state.quotes_data)} quote file(s)")
            for fn, qs in st.session_state.quotes_data.items():
                st.markdown(f"- **{fn}**: {len(qs)} items (${sum(q['Total_Price'] for q in qs):,.2f})")
        
        qf = st.file_uploader("Select Quotes", type=['pdf', 'csv', 'xlsx', 'xls'], key="quote", accept_multiple_files=True)
        if qf:
            for f in qf:
                if f.name not in st.session_state.quotes_data:
                    with st.spinner(f"Processing {f.name}..."):
                        parsed = parse_uploaded_file(f)
                        if parsed:
                            q = process_quote_file(parsed)
                            if q:
                                st.session_state.quotes_data[f.name] = q
                                st.success(f"✅ {len(q)} items from {f.name}")
        
        if st.session_state.quotes_data and st.button("🗑️ Clear Quotes"):
            st.session_state.quotes_data = {}
            st.rerun()
    
    st.markdown("---")
    if st.button("🔄 Reset All"):
        st.session_state.equipment_schedule = None
        st.session_state.quotes_data = {}
        st.session_state.drawing_filename = None
        st.rerun()

with tabs[1]:
    if not st.session_state.equipment_schedule:
        st.warning("⚠️ Upload drawing first")
    elif not st.session_state.quotes_data:
        st.warning("⚠️ Upload quotes first")
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
        col1.metric("💰 Total Quoted", f"${df['Total_Price'].sum():,.2f}")
        col2.metric("📦 Total Items", f"{len(df)} ({len(act)} actionable)")
        
        ch1, ch2 = st.columns(2)
        with ch1:
            vc = df['Status'].value_counts().reset_index()
            vc.columns = ['Status', 'Count']
            cm = {'✓ Quoted': '#28a745', '❌ MISSING': '#dc3545', '⚠ Qty Mismatch': '#ffc107', 
                  '⚠ Needs Install': '#fd7e14', 'IH Supply': '#6c757d', 'Existing': '#adb5bd', 'N/A': '#e9ecef'}
            fig = px.pie(vc, values='Count', names='Status', color='Status', color_discrete_map=cm)
            st.plotly_chart(fig, use_container_width=True)
        with ch2:
            cat_df = df[df['Supplier_Code'].notna()].groupby('Supplier_Code').size().reset_index(name='Items')
            cat_df['Label'] = cat_df['Supplier_Code'].astype(int).map(lambda x: f"Code {x}")
            fig2 = px.bar(cat_df, x='Label', y='Items', title="Items by Supplier Code")
            st.plotly_chart(fig2, use_container_width=True)

with tabs[2]:
    if st.session_state.equipment_schedule and st.session_state.quotes_data:
        all_q = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_schedule_vs_quotes(st.session_state.equipment_schedule, all_q)
        
        st.subheader("📋 Full Report")
        col1, col2 = st.columns(2)
        status_opts = df['Status'].unique().tolist()
        filt_status = col1.multiselect("Status", status_opts, default=status_opts)
        cat_opts = sorted([int(c) for c in df['Supplier_Code'].dropna().unique()])
        filt_cat = col2.multiselect("Supplier Code", cat_opts, default=cat_opts)
        
        fdf = df[df['Status'].isin(filt_status)]
        if filt_cat:
            fdf = fdf[(fdf['Supplier_Code'].isin(filt_cat)) | (fdf['Supplier_Code'].isna())]
        
        def hl(row):
            cm = {'✓ Quoted': 'background-color:#d4edda', '❌ MISSING': 'background-color:#f8d7da', 
                  '⚠ Qty Mismatch': 'background-color:#fff3cd', '⚠ Needs Install': 'background-color:#ffe5d0'}
            return [cm.get(row['Status'], '')] * len(row)
        st.dataframe(fdf.style.apply(hl, axis=1), height=450, use_container_width=True)
        
        st.subheader("🚨 Critical Missing (Codes 5 & 6)")
        crit = df[(df['Status'] == '❌ MISSING') & (df['Supplier_Code'].isin([5, 6]))]
        if not crit.empty:
            st.error(f"{len(crit)} critical items need quotes!")
            st.dataframe(crit[['No', 'Equip_Num', 'Description', 'Schedule_Qty', 'Supplier_Desc']], use_container_width=True)
        else:
            st.success("✅ No critical missing!")

with tabs[3]:
    if st.session_state.equipment_schedule and st.session_state.quotes_data:
        all_q = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_schedule_vs_quotes(st.session_state.equipment_schedule, all_q)
        
        st.subheader("🔢 By Supplier Code")
        summary = []
        for code, desc in SUPPLIER_CODES.items():
            ci = df[df['Supplier_Code'] == code]
            if len(ci):
                summary.append({'Code': code, 'Description': desc, 'Items': len(ci),
                    'Quoted': len(ci[ci['Status'] == '✓ Quoted']), 'Missing': len(ci[ci['Status'] == '❌ MISSING']),
                    'Value': f"${ci['Total_Price'].sum():,.2f}"})
        st.dataframe(pd.DataFrame(summary), use_container_width=True)
        
        no_cat = df[df['Supplier_Code'].isna()]
        if len(no_cat):
            st.subheader("📋 SPARE/Placeholder Items")
            st.dataframe(no_cat[['No', 'Equip_Num', 'Description', 'Schedule_Qty']], use_container_width=True)

with tabs[4]:
    st.subheader("📥 Export")
    if st.session_state.equipment_schedule:
        eq_df = pd.DataFrame(st.session_state.equipment_schedule)
        out_eq = io.BytesIO()
        eq_df.to_excel(out_eq, index=False)
        out_eq.seek(0)
        st.download_button("📥 Equipment Schedule", out_eq, f"Equipment_{datetime.now().strftime('%Y%m%d')}.xlsx")
    
    if st.session_state.equipment_schedule and st.session_state.quotes_data:
        all_q = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_schedule_vs_quotes(st.session_state.equipment_schedule, all_q)
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as w:
            df.to_excel(w, sheet_name='Analysis', index=False)
            df[df['Status'] == '❌ MISSING'].to_excel(w, sheet_name='Missing', index=False)
            df[(df['Status'] == '❌ MISSING') & (df['Supplier_Code'].isin([5, 6]))].to_excel(w, sheet_name='Critical', index=False)
            pd.DataFrame(all_q).to_excel(w, sheet_name='Quotes', index=False)
        out.seek(0)
        st.download_button("📥 Full Report", out, f"Analysis_{datetime.now().strftime('%Y%m%d')}.xlsx")

st.markdown("---")
st.markdown("<center>Equipment Quote Analyzer v6.2</center>", unsafe_allow_html=True)
