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

SKIP_WORDS = ['equipment list', 'electrical services', 'mechanical services', 
    'project', 'title', 'drawing', 'revision', 'zeidler', 'copyright', 
    'issued', 'date', 'kitchen equipment fixture', 'this plan', 'all services', 
    'electrical contractor', 'mechanical contractor', 'kitchen contractor',
    'upon completion', 'at this point', 'project address', 'supplier code',
    'interior health', 'autodesk docs', 'conn. type', 'elec. ri']

if 'equipment_schedule' not in st.session_state:
    st.session_state.equipment_schedule = None
if 'quotes_data' not in st.session_state:
    st.session_state.quotes_data = {}
if 'drawing_filename' not in st.session_state:
    st.session_state.drawing_filename = None

def parse_pdf_file(uploaded_file):
    if not PDF_SUPPORT:
        return None
    text_content = []
    uploaded_file.seek(0)
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                text_content.append(text)
    return "\n".join(text_content)

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
        text = parse_pdf_file(uploaded_file)
        return {'type': 'pdf', 'text': text, 'filename': uploaded_file.name}
    elif ext in ['xlsx', 'xls']:
        sheets = parse_excel_file(uploaded_file)
        return {'type': 'excel', 'sheets': sheets, 'filename': uploaded_file.name}
    elif ext == 'csv':
        sheets = parse_csv_file(uploaded_file)
        return {'type': 'csv', 'sheets': sheets, 'filename': uploaded_file.name}
    return None

def parse_equipment_from_text(text):
    equipment_list = []
    lines = text.split('\n')
    
    for line in lines:
        line = line.strip()
        if not line or len(line) < 8:
            continue
        
        line_lower = line.lower()
        skip = False
        for sw in SKIP_WORDS:
            if sw in line_lower:
                skip = True
                break
        if skip:
            continue
        
        if line.startswith(('No.', 'NEW ', '1 :', 'K-', '300,', 'T 403', 'E1 ', 'E2 ', 'M1 ', 'M2 ')):
            continue
        
        match = re.match(r'^(\d{1,2}[a-z]?)\s+(.+)$', line, re.IGNORECASE)
        if not match:
            continue
        
        item_no = match.group(1)
        rest = match.group(2).strip()
        
        try:
            num_only = re.sub(r'[a-z]', '', item_no, flags=re.IGNORECASE)
            if int(num_only) > 90:
                continue
        except:
            continue
        
        if re.match(r'^(IH SUPPLY|CONTRACTOR|EXISTING)', rest, re.IGNORECASE):
            continue
        
        equip_num = '-'
        eq_match = re.match(r'^(\d{4,}\.?\d*)\s+(.+)$', rest)
        if eq_match:
            equip_num = eq_match.group(1)
            rest = eq_match.group(2).strip()
        
        if rest.startswith('- '):
            equip_num = '-'
            rest = rest[2:].strip()
        
        elec_match = re.search(r'(\d+A\s+\d+V|\d+KW|JUNCTION|RECEPTACLE|SEE NOTE|SERVICES|LIGHTS|FFD|STUB-UP|WASTE TO)', rest, re.IGNORECASE)
        if elec_match:
            rest = rest[:elec_match.start()].strip()
        
        qty_cat = re.search(r'\s+(\d+)\s+([1-8])\s*$', rest)
        qty_dash = re.search(r'\s+(\d+)\s+(-)\s*$', rest)
        
        if qty_cat:
            description = rest[:qty_cat.start()].strip()
            qty = int(qty_cat.group(1))
            category = int(qty_cat.group(2))
        elif qty_dash:
            description = rest[:qty_dash.start()].strip()
            qty = int(qty_dash.group(1))
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
            'New_Equip_Num': equip_num,
            'Description': description,
            'Qty': qty,
            'Category': category
        })
    
    return equipment_list

def extract_equipment_from_dataframe(df):
    equipment_list = []
    df = df.dropna(how='all').reset_index(drop=True)
    df.columns = df.columns.astype(str).str.strip().str.lower()
    
    col_map = {
        'no': ['no', 'no.', 'item', 'item #', 'item no', 'number', '#'],
        'new_equip_num': ['new equipment number', 'equipment number', 'equip num', 'equip no'],
        'description': ['description', 'desc', 'equipment', 'name'],
        'qty': ['qty', 'qty.', 'quantity', 'count'],
        'category': ['category', 'cat', 'supplier code', 'code']
    }
    
    found = {}
    for key, opts in col_map.items():
        for col in df.columns:
            col_clean = col.lower().strip()
            if col_clean in opts:
                found[key] = col
                break
            for o in opts:
                if o in col_clean:
                    found[key] = col
                    break
    
    if 'no' not in found or 'description' not in found:
        if len(df.columns) >= 3:
            found['no'] = df.columns[0]
            found['new_equip_num'] = df.columns[1]
            found['description'] = df.columns[2]
    
    if 'no' not in found or 'description' not in found:
        return None
    
    for idx, row in df.iterrows():
        try:
            no_val = str(row.get(found['no'], '')).strip()
            desc_val = str(row.get(found['description'], '')).strip()
            
            if not no_val or no_val.lower() in ('nan', '', 'no', 'no.', 'item'):
                continue
            if not desc_val or desc_val.lower() in ('nan', '', 'description'):
                continue
            
            equip_num = '-'
            if 'new_equip_num' in found:
                en = str(row.get(found['new_equip_num'], '')).strip()
                if en and en.lower() not in ('nan', ''):
                    equip_num = en
            
            qty = 1
            if 'qty' in found:
                try:
                    qv = str(row.get(found['qty'], 1))
                    qv = re.sub(r'[^\d.]', '', qv)
                    if qv:
                        qty = int(float(qv))
                except:
                    pass
            
            cat = None
            if 'category' in found:
                try:
                    cv = str(row.get(found['category'], ''))
                    cv = re.sub(r'[^\d]', '', cv)
                    if cv:
                        cat = int(float(cv))
                except:
                    pass
            
            equipment_list.append({
                'No': no_val,
                'New_Equip_Num': equip_num,
                'Description': desc_val,
                'Qty': qty,
                'Category': cat
            })
        except:
            continue
    
    return equipment_list if equipment_list else None

def process_drawing_file(parsed_data):
    equipment_list = []
    
    if parsed_data['type'] == 'pdf' and parsed_data.get('text'):
        equipment_list = parse_equipment_from_text(parsed_data['text'])
    elif parsed_data['type'] in ['excel', 'csv'] and parsed_data.get('sheets'):
        for df in parsed_data['sheets'].values():
            ext = extract_equipment_from_dataframe(df)
            if ext:
                equipment_list.extend(ext)
    
    seen = set()
    unique = []
    for item in equipment_list:
        if item['No'] not in seen:
            seen.add(item['No'])
            unique.append(item)
    return unique

def parse_quote_from_text(text, filename):
    quotes = []
    current_item = None
    current_desc = None
    current_qty = 0
    current_unit = 0
    current_total = 0
    
    skip_words = ['page ', 'canadian restaurant', 'bird construc', 'fwg ltc', 
                  'item qty description', 'sell total', 'merchandise', 'prices are in', 'quote valid']
    
    for line in text.split('\n'):
        line = line.strip()
        if not line:
            continue
        
        ll = line.lower()
        skip = False
        for s in skip_words:
            if s in ll:
                skip = True
                break
        if skip:
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
                    'Item_No': current_item, 
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
            current_unit = float(prices[0].replace(',', '')) if prices else 0
            current_total = float(prices[1].replace(',', '')) if len(prices) >= 2 else current_unit * current_qty
    
    if current_item and current_desc:
        quotes.append({
            'Item_No': current_item, 
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
        'item': ['item', 'item no', 'no', 'no.'], 
        'description': ['description', 'desc', 'equipment', 'name'],
        'qty': ['qty', 'quantity'], 
        'unit_price': ['sell', 'unit price', 'price'], 
        'total_price': ['sell total', 'total', 'total price', 'amount']
    }
    
    found = {}
    for k, opts in col_map.items():
        for c in df.columns:
            if c.lower().strip() in opts:
                found[k] = c
                break
    
    for idx, row in df.iterrows():
        try:
            desc = str(row.get(found.get('description', ''), '')).strip()
            if not desc or desc.lower() in ('nan', '', 'nic'):
                continue
            
            item = str(row.get(found.get('item', ''), '')).strip()
            
            qty = 1
            if 'qty' in found:
                try:
                    qv = str(row.get(found['qty'], 1)).replace('ea', '').replace(',', '').strip()
                    if qv:
                        qty = int(float(qv))
                except:
                    pass
            
            up = 0
            if 'unit_price' in found:
                try:
                    uv = str(row.get(found['unit_price'], 0)).replace('$', '').replace(',', '').strip()
                    if uv:
                        up = float(uv)
                except:
                    pass
            
            tp = 0
            if 'total_price' in found:
                try:
                    tv = str(row.get(found['total_price'], 0)).replace('$', '').replace(',', '').strip()
                    if tv:
                        tp = float(tv)
                except:
                    pass
            
            if tp == 0:
                tp = up * qty
            
            quotes.append({
                'Item_No': item, 
                'Description': desc, 
                'Qty': qty, 
                'Unit_Price': up, 
                'Total_Price': tp, 
                'Source_File': filename
            })
        except:
            continue
    return quotes

def process_quote_file(parsed_data):
    quotes = []
    if parsed_data['type'] == 'pdf' and parsed_data.get('text'):
        quotes = parse_quote_from_text(parsed_data['text'], parsed_data['filename'])
    elif parsed_data['type'] in ['excel', 'csv'] and parsed_data.get('sheets'):
        for df in parsed_data['sheets'].values():
            quotes.extend(extract_quotes_from_dataframe(df, parsed_data['filename']))
    
    seen = set()
    unique = []
    for q in quotes:
        if q['Item_No'] and q['Item_No'] not in seen:
            seen.add(q['Item_No'])
            unique.append(q)
    return unique

def match_quote_to_drawing(drawing_item, quotes):
    drawing_no = str(drawing_item['No']).strip().lower()
    
    for q in quotes:
        if str(q.get('Item_No', '')).strip().lower() == drawing_no:
            return q
    
    try:
        drawing_num = int(re.sub(r'[a-zA-Z]', '', drawing_no))
        for q in quotes:
            try:
                quote_num = int(re.sub(r'[a-zA-Z]', '', str(q.get('Item_No', '')).strip()))
                if drawing_num == quote_num:
                    return q
            except:
                pass
    except:
        pass
    return None

def analyze_drawing_vs_quotes(drawing_schedule, quotes):
    analysis = []
    
    for item in drawing_schedule:
        match = match_quote_to_drawing(item, quotes)
        cat = item.get('Category')
        
        if cat in [1, 2, 3]:
            status = "IH Supply"
            issue = "IH handles supply and install"
        elif cat == 8:
            status = "Existing"
            issue = "Existing or relocated"
        elif cat is None or item.get('Description', '').upper() == 'SPARE':
            status = "N/A"
            issue = "Spare or placeholder"
        elif match:
            if match['Qty'] == item['Qty']:
                status = "Quoted"
                issue = None
            else:
                status = "Qty Mismatch"
                issue = "Drawing: " + str(item['Qty']) + " Quote: " + str(match['Qty'])
        else:
            if cat == 7:
                status = "Needs Install"
                issue = "IH supplies - needs install pricing"
            elif cat in [5, 6]:
                status = "MISSING"
                issue = "Critical - requires quote"
            else:
                status = "MISSING"
                issue = "Not found in quotes"
        
        analysis.append({
            'Drawing_No': item['No'],
            'New_Equip_Num': item.get('New_Equip_Num', '-'),
            'Description': item['Description'],
            'Drawing_Qty': item['Qty'],
            'Category': cat,
            'Category_Desc': SUPPLIER_CODES.get(cat, 'N/A') if cat else 'N/A',
            'Quote_Item_No': match['Item_No'] if match else '-',
            'Quote_Qty': match['Qty'] if match else 0,
            'Unit_Price': match['Unit_Price'] if match else 0,
            'Total_Price': match['Total_Price'] if match else 0,
            'Quote_Source': match['Source_File'] if match else '-',
            'Status': status,
            'Issue': issue
        })
    return pd.DataFrame(analysis)

# UI
st.markdown("## Drawing vs Quote Analyzer")

if not PDF_SUPPORT:
    st.warning("Install pdfplumber: pip install pdfplumber")

tabs = st.tabs(["Upload", "Dashboard", "Analysis", "Summary", "Export"])

with tabs[0]:
    c1, c2 = st.columns(2)
    
    with c1:
        st.subheader("Drawing")
        if st.session_state.equipment_schedule:
            count = len(st.session_state.equipment_schedule)
            st.success(st.session_state.drawing_filename + " (" + str(count) + " items)")
            with st.expander("View Equipment"):
                st.dataframe(pd.DataFrame(st.session_state.equipment_schedule), height=250)
        
        df_file = st.file_uploader("Upload Drawing", type=['pdf', 'csv', 'xlsx', 'xls'], key="draw")
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
                        st.error("Could not extract equipment")
    
    with c2:
        st.subheader("Quotations")
        if st.session_state.quotes_data:
            st.success(str(len(st.session_state.quotes_data)) + " quote file(s)")
            for fn, qs in st.session_state.quotes_data.items():
                total = sum(q['Total_Price'] for q in qs)
                st.write(fn + ": " + str(len(qs)) + " items ($" + "{:,.2f}".format(total) + ")")
        
        qf = st.file_uploader("Upload Quotes", type=['pdf', 'csv', 'xlsx', 'xls'], key="quote", accept_multiple_files=True)
        if qf:
            for f in qf:
                if f.name not in st.session_state.quotes_data:
                    with st.spinner("Processing..."):
                        parsed = parse_uploaded_file(f)
                        if parsed:
                            q = process_quote_file(parsed)
                            if q:
                                st.session_state.quotes_data[f.name] = q
                                st.success(str(len(q)) + " items from " + f.name)
        
        if st.session_state.quotes_data:
            if st.button("Clear Quotes"):
                st.session_state.quotes_data = {}
                st.rerun()
    
    st.markdown("---")
    if st.button("Reset All"):
        st.session_state.equipment_schedule = None
        st.session_state.quotes_data = {}
        st.session_state.drawing_filename = None
        st.rerun()

with tabs[1]:
    if not st.session_state.equipment_schedule:
        st.warning("Upload drawing first")
    elif not st.session_state.quotes_data:
        st.warning("Upload quotes first")
    else:
        all_q = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_drawing_vs_quotes(st.session_state.equipment_schedule, all_q)
        
        st.subheader("Coverage Summary")
        act = df[~df['Status'].isin(['IH Supply', 'Existing', 'N/A'])]
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Quoted", len(act[act['Status'] == 'Quoted']))
        c2.metric("Missing", len(act[act['Status'] == 'MISSING']))
        c3.metric("Qty Mismatch", len(act[act['Status'] == 'Qty Mismatch']))
        c4.metric("Needs Install", len(act[act['Status'] == 'Needs Install']))
        
        col1, col2 = st.columns(2)
        col1.metric("Total Quoted", "$" + "{:,.2f}".format(df['Total_Price'].sum()))
        col2.metric("Total Items", str(len(df)) + " (" + str(len(act)) + " actionable)")
        
        ch1, ch2 = st.columns(2)
        with ch1:
            vc = df['Status'].value_counts().reset_index()
            vc.columns = ['Status', 'Count']
            colors = {'Quoted': '#28a745', 'MISSING': '#dc3545', 'Qty Mismatch': '#ffc107', 
                      'Needs Install': '#fd7e14', 'IH Supply': '#6c757d', 'Existing': '#adb5bd', 'N/A': '#e9ecef'}
            fig = px.pie(vc, values='Count', names='Status', color='Status', color_discrete_map=colors)
            st.plotly_chart(fig, use_container_width=True)
        
        with ch2:
            cat_df = df[df['Category'].notna()].groupby('Category').size().reset_index(name='Items')
            cat_df['Label'] = cat_df['Category'].astype(int).apply(lambda x: "Cat " + str(x))
            fig2 = px.bar(cat_df, x='Label', y='Items', title="By Category")
            st.plotly_chart(fig2, use_container_width=True)

with tabs[2]:
    if st.session_state.equipment_schedule and st.session_state.quotes_data:
        all_q = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_drawing_vs_quotes(st.session_state.equipment_schedule, all_q)
        
        st.subheader("Full Analysis")
        
        col1, col2 = st.columns(2)
        status_opts = df['Status'].unique().tolist()
        filt_status = col1.multiselect("Status", status_opts, default=status_opts)
        cat_opts = sorted([int(c) for c in df['Category'].dropna().unique()])
        filt_cat = col2.multiselect("Category", cat_opts, default=cat_opts)
        
        fdf = df[df['Status'].isin(filt_status)]
        if filt_cat:
            fdf = fdf[(fdf['Category'].isin(filt_cat)) | (fdf['Category'].isna())]
        
        def highlight(row):
            colors = {
                'Quoted': 'background-color:#d4edda', 
                'MISSING': 'background-color:#f8d7da', 
                'Qty Mismatch': 'background-color:#fff3cd', 
                'Needs Install': 'background-color:#ffe5d0'
            }
            return [colors.get(row['Status'], '')] * len(row)
        
        cols = ['Drawing_No', 'New_Equip_Num', 'Description', 'Drawing_Qty', 'Category', 
                'Quote_Item_No', 'Quote_Qty', 'Unit_Price', 'Total_Price', 'Status', 'Issue']
        st.dataframe(fdf[cols].style.apply(highlight, axis=1), height=450, use_container_width=True)
        
        st.subheader("Critical Missing (Cat 5 & 6)")
        crit = df[(df['Status'] == 'MISSING') & (df['Category'].isin([5, 6]))]
        if len(crit) > 0:
            st.error(str(len(crit)) + " items need quotes!")
            st.dataframe(crit[['Drawing_No', 'New_Equip_Num', 'Description', 'Drawing_Qty', 'Category_Desc']])
        else:
            st.success("No critical missing!")

with tabs[3]:
    if st.session_state.equipment_schedule and st.session_state.quotes_data:
        all_q = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_drawing_vs_quotes(st.session_state.equipment_schedule, all_q)
        
        st.subheader("By Category")
        summary = []
        for code, desc in SUPPLIER_CODES.items():
            ci = df[df['Category'] == code]
            if len(ci) > 0:
                summary.append({
                    'Cat': code, 
                    'Description': desc, 
                    'Items': len(ci),
                    'Quoted': len(ci[ci['Status'] == 'Quoted']), 
                    'Missing': len(ci[ci['Status'] == 'MISSING']),
                    'Value': "$" + "{:,.2f}".format(ci['Total_Price'].sum())
                })
        st.dataframe(pd.DataFrame(summary), use_container_width=True)
        
        no_cat = df[df['Category'].isna()]
        if len(no_cat) > 0:
            st.subheader("SPARE Items")
            st.dataframe(no_cat[['Drawing_No', 'New_Equip_Num', 'Description', 'Drawing_Qty']])

with tabs[4]:
    st.subheader("Export")
    
    if st.session_state.equipment_schedule:
        eq_df = pd.DataFrame(st.session_state.equipment_schedule)
        out_eq = io.BytesIO()
        eq_df.to_excel(out_eq, index=False)
        out_eq.seek(0)
        fname = "Drawing_" + datetime.now().strftime('%Y%m%d') + ".xlsx"
        st.download_button("Download Drawing", out_eq, fname)
    
    if st.session_state.equipment_schedule and st.session_state.quotes_data:
        all_q = [q for qs in st.session_state.quotes_data.values() for q in qs]
        df = analyze_drawing_vs_quotes(st.session_state.equipment_schedule, all_q)
        
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as w:
            df.to_excel(w, sheet_name='Analysis', index=False)
            df[df['Status'] == 'MISSING'].to_excel(w, sheet_name='Missing', index=False)
            df[(df['Status'] == 'MISSING') & (df['Category'].isin([5, 6]))].to_excel(w, sheet_name='Critical', index=False)
            pd.DataFrame(all_q).to_excel(w, sheet_name='Quotes', index=False)
        out.seek(0)
        fname = "Analysis_" + datetime.now().strftime('%Y%m%d') + ".xlsx"
        st.download_button("Download Analysis", out, fname)

st.markdown("---")
st.write("Drawing Quote Analyzer v6.5")
