import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
from datetime import datetime

st.set_page_config(page_title="Kitchen Quote Analyzer", page_icon="🍳", layout="wide")

st.markdown("""
<style>
    .main-header { font-size: 2.5rem; font-weight: bold; color: #1f4e79; }
    .sub-header { font-size: 1.2rem; color: #666; }
    div[data-testid="stMetricValue"] { font-size: 1.8rem; }
    .upload-section { background-color: #f8f9fa; padding: 1rem; border-radius: 8px; margin: 0.5rem 0; }
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

# Session state initialization
if 'schedule_df' not in st.session_state:
    st.session_state.schedule_df = None
if 'quotes' not in st.session_state:
    st.session_state.quotes = {}
if 'project_name' not in st.session_state:
    st.session_state.project_name = "My Project"

def create_schedule_template():
    """Create downloadable template for equipment schedule"""
    template = pd.DataFrame({
        'Item': ['1', '2', '3', '1a'],
        'Description': ['STORAGE SHELVING', 'WALK-IN FREEZER', 'EVAPORATOR COIL', 'STORAGE SHELVING (EXISTING)'],
        'Qty': [3, 1, 1, 2],
        'Category': [1, 5, 5, 8]
    })
    return template

def create_quote_template():
    """Create downloadable template for vendor quotes"""
    template = pd.DataFrame({
        'Item': ['2', '3', '10'],
        'Description': ['WALK-IN FREEZER c/w INSULATED FLOOR', 'EVAPORATOR COIL', 'CONDENSING UNIT STAND'],
        'Qty': [1, 1, 2],
        'Unit_Price': [97980.27, 70727.09, 2206.08],
        'Total': [97980.27, 70727.09, 4412.16]
    })
    return template

def parse_schedule_file(uploaded_file):
    """Parse uploaded schedule file (CSV or Excel)"""
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        # Normalize column names
        df.columns = df.columns.str.strip().str.title()
        
        # Map common column variations
        col_map = {
            'Item #': 'Item', 'Item No': 'Item', 'Item Number': 'Item', 'No': 'Item', 'No.': 'Item',
            'Desc': 'Description', 'Equipment': 'Description', 'Name': 'Description',
            'Quantity': 'Qty', 'Qty.': 'Qty', 'Count': 'Qty',
            'Cat': 'Category', 'Code': 'Category', 'Supplier Code': 'Category', 'Type': 'Category'
        }
        df = df.rename(columns={k: v for k, v in col_map.items() if k in df.columns})
        
        required = ['Item', 'Description', 'Qty']
        missing = [c for c in required if c not in df.columns]
        if missing:
            return None, f"Missing required columns: {', '.join(missing)}"
        
        df['Item'] = df['Item'].astype(str).str.strip()
        df['Qty'] = pd.to_numeric(df['Qty'], errors='coerce').fillna(0).astype(int)
        
        if 'Category' not in df.columns:
            df['Category'] = None
        else:
            df['Category'] = pd.to_numeric(df['Category'], errors='coerce')
        
        return df, None
    except Exception as e:
        return None, str(e)

def parse_quote_file(uploaded_file, vendor_name=""):
    """Parse uploaded quote file (CSV or Excel)"""
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        df.columns = df.columns.str.strip().str.title()
        
        col_map = {
            'Item #': 'Item', 'Item No': 'Item', 'Item Number': 'Item', 'No': 'Item', 'No.': 'Item',
            'Desc': 'Description', 'Equipment': 'Description', 'Name': 'Description',
            'Quantity': 'Qty', 'Qty.': 'Qty', 'Count': 'Qty',
            'Unit Price': 'Unit_Price', 'Price': 'Unit_Price', 'Unit': 'Unit_Price', 'Unit Cost': 'Unit_Price',
            'Extended': 'Total', 'Ext Price': 'Total', 'Amount': 'Total', 'Line Total': 'Total', 'Ext': 'Total'
        }
        df = df.rename(columns={k: v for k, v in col_map.items() if k in df.columns})
        
        required = ['Item', 'Qty']
        missing = [c for c in required if c not in df.columns]
        if missing:
            return None, f"Missing required columns: {', '.join(missing)}"
        
        df['Item'] = df['Item'].astype(str).str.strip()
        df['Qty'] = pd.to_numeric(df['Qty'], errors='coerce').fillna(0).astype(int)
        
        if 'Description' not in df.columns:
            df['Description'] = ''
        
        if 'Unit_Price' not in df.columns:
            df['Unit_Price'] = 0
        else:
            df['Unit_Price'] = pd.to_numeric(df['Unit_Price'], errors='coerce').fillna(0)
        
        if 'Total' not in df.columns:
            df['Total'] = df['Qty'] * df['Unit_Price']
        else:
            df['Total'] = pd.to_numeric(df['Total'], errors='coerce').fillna(0)
        
        quote_name = vendor_name if vendor_name else uploaded_file.name.replace('.csv', '').replace('.xlsx', '').replace('.xls', '')
        
        quote_data = {
            "vendor": vendor_name if vendor_name else "Unknown Vendor",
            "date": datetime.now().strftime("%Y-%m-%d"),
            "items": df[['Item', 'Description', 'Qty', 'Unit_Price', 'Total']].to_dict('records'),
            "subtotal": df['Total'].sum(),
            "gst": df['Total'].sum() * 0.05,
            "pst": 0,
            "total": df['Total'].sum() * 1.05
        }
        return {quote_name: quote_data}, None
    except Exception as e:
        return None, str(e)

def get_supplier_code_name(code):
    return SUPPLIER_CODES.get(code, "Unknown") if pd.notna(code) else "N/A"

def is_ih_supply_install(category):
    return category in [1, 2, 3]

def analyze_quotes(schedule_df, quotes):
    results = []
    for _, row in schedule_df.iterrows():
        item_num = str(row['Item']).strip()
        if 'SPARE' in str(row['Description']).upper():
            continue
        
        quote_info = {"found": False, "quote_name": None, "qty": 0, "unit_price": 0, "total": 0}
        for quote_name, quote_data in quotes.items():
            for q_item in quote_data.get("items", []):
                q_item_num = str(q_item["Item"]).strip()
                if q_item_num == item_num:
                    quote_info = {
                        "found": True,
                        "quote_name": quote_name,
                        "qty": q_item["Qty"],
                        "unit_price": q_item["Unit_Price"],
                        "total": q_item["Total"]
                    }
                    break
            if quote_info["found"]:
                break
        
        cat = row.get('Category')
        if pd.isna(cat) or cat is None:
            status, issue = "N/A", "Spare/Undefined"
        elif is_ih_supply_install(cat):
            status, issue = "IH Supply", "IH Supply & Install - Excluded"
        elif quote_info["found"]:
            if quote_info["qty"] == row['Qty']:
                status, issue = "✓ Included", "None"
            else:
                status, issue = "⚠ Qty Mismatch", f"Expected {row['Qty']}, got {quote_info['qty']}"
        else:
            status, issue = "✗ Missing", "Not found in any quote"
        
        results.append({
            "Item #": item_num, "Description": row['Description'], "Schedule Qty": row['Qty'],
            "Quote Qty": quote_info["qty"] if quote_info["found"] else "-",
            "Supplier Code": cat,
            "Supplier Code Name": get_supplier_code_name(cat),
            "Unit Price": quote_info['unit_price'] if quote_info["found"] else 0,
            "Total Price": quote_info['total'] if quote_info["found"] else 0,
            "Found in Quote": quote_info["quote_name"] if quote_info["found"] else "-",
            "Status": status, "Issue": issue
        })
    return pd.DataFrame(results)

def get_supplier_code_summary(schedule_df, results_df):
    summary_data = []
    for code in range(1, 9):
        schedule_items = schedule_df[schedule_df['Category'] == code]
        line_items = len(schedule_items)
        total_qty = schedule_items['Qty'].sum() if not schedule_items.empty else 0
        
        code_results = results_df[results_df['Supplier Code'] == code]
        quoted = len(code_results[code_results['Status'] == '✓ Included'])
        missing = len(code_results[code_results['Status'] == '✗ Missing'])
        mismatch = len(code_results[code_results['Status'] == '⚠ Qty Mismatch'])
        quoted_val = code_results[code_results['Status'].isin(['✓ Included', '⚠ Qty Mismatch'])]['Total Price'].sum()
        
        coverage = f"{(quoted / max(line_items, 1)) * 100:.1f}%" if code not in [1, 2, 3] and line_items > 0 else "N/A"
        
        summary_data.append({
            "Supplier Code": code, "Description": SUPPLIER_CODES[code],
            "Schedule Line Items": line_items, "Schedule Total Qty": int(total_qty),
            "Quoted Items": quoted, "Missing Items": missing, "Qty Mismatch": mismatch,
            "Quoted Value": quoted_val, "Quote Required": "No" if code in [1, 2, 3] else "Yes",
            "Coverage %": coverage
        })
    return pd.DataFrame(summary_data)

def create_excel_report(schedule_df, results_df, quotes, supplier_summary_df, project_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        summary = {
            "Metric": ["Project", "Report Date", "Total Items", "Quoted", "Missing", "Mismatch", "IH Excluded", "Quoted Value"],
            "Value": [project_name, datetime.now().strftime("%Y-%m-%d %H:%M"),
                     len(results_df), len(results_df[results_df['Status'] == '✓ Included']),
                     len(results_df[results_df['Status'] == '✗ Missing']),
                     len(results_df[results_df['Status'] == '⚠ Qty Mismatch']),
                     len(results_df[results_df['Status'] == 'IH Supply']),
                     f"${results_df['Total Price'].sum():,.2f}"]}
        pd.DataFrame(summary).to_excel(writer, sheet_name='Executive Summary', index=False)
        
        sup_disp = supplier_summary_df.copy()
        sup_disp['Quoted Value'] = sup_disp['Quoted Value'].apply(lambda x: f"${x:,.2f}")
        sup_disp.to_excel(writer, sheet_name='Supplier Code Summary', index=False)
        
        analysis = results_df.copy()
        analysis['Unit Price'] = analysis['Unit Price'].apply(lambda x: f"${x:,.2f}" if x > 0 else "-")
        analysis['Total Price'] = analysis['Total Price'].apply(lambda x: f"${x:,.2f}" if x > 0 else "-")
        analysis.to_excel(writer, sheet_name='Full Analysis', index=False)
        
        results_df[results_df['Status'] == '✗ Missing'].to_excel(writer, sheet_name='Missing Items', index=False)
        results_df[results_df['Status'] == '⚠ Qty Mismatch'].to_excel(writer, sheet_name='Qty Mismatch', index=False)
        results_df[results_df['Status'] == '✓ Included'].to_excel(writer, sheet_name='Included Items', index=False)
        
        sched = schedule_df.copy()
        sched['Supplier Code Name'] = sched['Category'].apply(lambda x: SUPPLIER_CODES.get(x, 'N/A') if pd.notna(x) else 'N/A')
        sched.to_excel(writer, sheet_name='Complete Schedule', index=False)
        
        quote_list = []
        for qn, qd in quotes.items():
            for item in qd.get('items', []):
                quote_list.append({"Quote": qn, "Vendor": qd.get('vendor', 'N/A'), "Item #": item['Item'],
                                  "Description": item['Description'], "Qty": item['Qty'],
                                  "Unit Price": f"${item['Unit_Price']:,.2f}", "Total": f"${item['Total']:,.2f}"})
        if quote_list:
            pd.DataFrame(quote_list).to_excel(writer, sheet_name='Quote Details', index=False)
    output.seek(0)
    return output

# ============== MAIN APP ==============
st.markdown('<p class="main-header">🍳 Kitchen Equipment Quote Analyzer</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Upload equipment schedules from drawings and vendor quotes to analyze coverage</p>', unsafe_allow_html=True)

tabs = st.tabs(["📤 Upload & Config", "📊 Dashboard", "📋 Detailed Analysis", "🔢 Supplier Summary", "💰 Quote Details", "📥 Export"])

with tabs[0]:
    st.header("📤 Upload Files & Configuration")
    
    col1, col2 = st.columns(2)
    with col1:
        st.session_state.project_name = st.text_input("Project Name", value=st.session_state.project_name)
    
    st.markdown("---")
    
    # Equipment Schedule Upload
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📐 Equipment Schedule (from Drawings)")
        st.markdown('<div class="upload-section">', unsafe_allow_html=True)
        
        schedule_file = st.file_uploader("Upload Schedule (CSV or Excel)", type=['csv', 'xlsx', 'xls'], key="schedule")
        
        if schedule_file:
            df, error = parse_schedule_file(schedule_file)
            if error:
                st.error(f"Error parsing schedule: {error}")
            else:
                st.session_state.schedule_df = df
                st.success(f"✅ Loaded {len(df)} items from {schedule_file.name}")
        
        if st.session_state.schedule_df is not None:
            st.markdown(f"**Current Schedule:** {len(st.session_state.schedule_df)} items")
            with st.expander("Preview Schedule"):
                st.dataframe(st.session_state.schedule_df.head(10), use_container_width=True, hide_index=True)
        
        st.markdown("</div>", unsafe_allow_html=True)
        
        # Template download
        st.markdown("**Download Template:**")
        template_df = create_schedule_template()
        csv = template_df.to_csv(index=False)
        st.download_button("📄 Schedule Template", csv, "schedule_template.csv", "text/csv", key="sched_tmpl")
    
    with col2:
        st.subheader("💵 Vendor Quotes")
        st.markdown('<div class="upload-section">', unsafe_allow_html=True)
        
        vendor_name = st.text_input("Vendor Name (optional)", placeholder="e.g., Canadian Restaurant Supply")
        quote_file = st.file_uploader("Upload Quote (CSV or Excel)", type=['csv', 'xlsx', 'xls'], key="quote")
        
        if quote_file:
            quote_data, error = parse_quote_file(quote_file, vendor_name)
            if error:
                st.error(f"Error parsing quote: {error}")
            else:
                st.session_state.quotes.update(quote_data)
                st.success(f"✅ Added quote: {list(quote_data.keys())[0]}")
        
        if st.session_state.quotes:
            st.markdown(f"**Loaded Quotes:** {len(st.session_state.quotes)}")
            for qname, qdata in st.session_state.quotes.items():
                col_a, col_b = st.columns([3, 1])
                col_a.markdown(f"• {qname} ({len(qdata['items'])} items, ${qdata['subtotal']:,.2f})")
                if col_b.button("🗑️", key=f"del_{qname}"):
                    del st.session_state.quotes[qname]
                    st.rerun()
        
        st.markdown("</div>", unsafe_allow_html=True)
        
        st.markdown("**Download Template:**")
        q_template = create_quote_template()
        q_csv = q_template.to_csv(index=False)
        st.download_button("📄 Quote Template", q_csv, "quote_template.csv", "text/csv", key="quote_tmpl")
    
    st.markdown("---")
    
    # Supplier Code Reference
    st.subheader("📋 Supplier Code Reference")
    st.info("Use these codes in your schedule's 'Category' column to classify equipment by supply/install responsibility")
    ref_df = pd.DataFrame([{"Code": k, "Description": v, "Quote Required": "No" if k in [1,2,3] else "Yes"} for k,v in SUPPLIER_CODES.items()])
    st.dataframe(ref_df, use_container_width=True, hide_index=True)
    
    # Clear data buttons
    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    if col1.button("🗑️ Clear Schedule", use_container_width=True):
        st.session_state.schedule_df = None
        st.rerun()
    if col2.button("🗑️ Clear All Quotes", use_container_width=True):
        st.session_state.quotes = {}
        st.rerun()
    if col3.button("🔄 Reset All", use_container_width=True):
        st.session_state.schedule_df = None
        st.session_state.quotes = {}
        st.rerun()

# Check if we have data to analyze
has_data = st.session_state.schedule_df is not None and len(st.session_state.quotes) > 0

if has_data:
    results_df = analyze_quotes(st.session_state.schedule_df, st.session_state.quotes)
    supplier_summary_df = get_supplier_code_summary(st.session_state.schedule_df, results_df)
else:
    results_df = pd.DataFrame()
    supplier_summary_df = pd.DataFrame()

with tabs[1]:
    st.header("📊 Analysis Dashboard")
    
    if not has_data:
        st.warning("⚠️ Please upload an equipment schedule and at least one quote in the Upload tab to see analysis.")
    else:
        col1, col2, col3, col4, col5 = st.columns(5)
        total = len(results_df)
        included = len(results_df[results_df['Status'] == '✓ Included'])
        missing = len(results_df[results_df['Status'] == '✗ Missing'])
        mismatch = len(results_df[results_df['Status'] == '⚠ Qty Mismatch'])
        ih_exc = len(results_df[results_df['Status'] == 'IH Supply'])
        
        col1.metric("Total Items", total)
        col2.metric("✓ Included", included)
        col3.metric("✗ Missing", missing)
        col4.metric("⚠ Mismatch", mismatch)
        col5.metric("IH Excluded", ih_exc)
        
        st.markdown("---")
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📈 Coverage Status")
            status_counts = results_df['Status'].value_counts()
            colors = {'✓ Included': '#28a745', '✗ Missing': '#dc3545', '⚠ Qty Mismatch': '#ffc107', 'IH Supply': '#6c757d', 'N/A': '#adb5bd'}
            fig = px.pie(values=status_counts.values, names=status_counts.index, color=status_counts.index, color_discrete_map=colors, hole=0.4)
            fig.update_layout(height=350)
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            st.subheader("📊 Items by Supplier Code")
            code_counts = results_df[results_df['Supplier Code'].notna()].groupby('Supplier Code').size().reset_index(name='Count')
            if not code_counts.empty:
                code_counts['Supplier Code'] = code_counts['Supplier Code'].astype(int)
                fig2 = px.bar(code_counts, x='Supplier Code', y='Count', color='Supplier Code', text='Count')
                fig2.update_layout(height=350, showlegend=False)
                fig2.update_traces(textposition='outside')
                st.plotly_chart(fig2, use_container_width=True)
        
        st.markdown("---")
        st.subheader("💰 Financial Summary")
        col1, col2, col3 = st.columns(3)
        col1.metric("Quoted Equipment Value", f"${results_df['Total Price'].sum():,.2f}")
        col2.metric("Items Needing Quotes", missing)
        col3.metric("Total Quote Value", f"${sum(q.get('total',0) for q in st.session_state.quotes.values()):,.2f}")

with tabs[2]:
    st.header("📋 Detailed Quote vs Schedule Comparison")
    
    if not has_data:
        st.warning("⚠️ Please upload data in the Upload tab first.")
    else:
        col1, col2, col3, col4 = st.columns(4)
        show_ih = col1.checkbox("Show IH Supply", False)
        show_inc = col2.checkbox("Show Included", True)
        show_miss = col3.checkbox("Show Missing", True)
        show_mis = col4.checkbox("Show Mismatch", True)
        
        status_filter = []
        if show_inc: status_filter.append('✓ Included')
        if show_miss: status_filter.append('✗ Missing')
        if show_mis: status_filter.append('⚠ Qty Mismatch')
        if show_ih: status_filter.append('IH Supply')
        
        filtered = results_df[results_df['Status'].isin(status_filter)] if status_filter else results_df
        
        codes = st.multiselect("Filter by Supplier Code", list(range(1,9)), default=[4,5,6,7,8],
                              format_func=lambda x: f"{x}: {SUPPLIER_CODES[x][:25]}...")
        if codes:
            filtered = filtered[filtered['Supplier Code'].isin(codes)]
        
        st.markdown(f"### Results ({len(filtered)} items)")
        
        display = filtered.copy()
        display['Unit Price'] = display['Unit Price'].apply(lambda x: f"${x:,.2f}" if x > 0 else "-")
        display['Total Price'] = display['Total Price'].apply(lambda x: f"${x:,.2f}" if x > 0 else "-")
        
        def color_rows(row):
            c = {'✓ Included': '#d4edda', '✗ Missing': '#f8d7da', '⚠ Qty Mismatch': '#fff3cd', 'IH Supply': '#e2e3e5'}
            return [f'background-color: {c.get(row["Status"], "")}'] * len(row)
        
        st.dataframe(display.style.apply(color_rows, axis=1), use_container_width=True, hide_index=True, height=500)

with tabs[3]:
    st.header("🔢 Supplier Code Summary")
    
    if not has_data:
        st.warning("⚠️ Please upload data in the Upload tab first.")
    else:
        disp_sum = supplier_summary_df.copy()
        disp_sum['Quoted Value'] = disp_sum['Quoted Value'].apply(lambda x: f"${x:,.2f}")
        
        def color_sum(row):
            if row['Quote Required'] == 'No': return ['background-color: #e2e3e5'] * len(row)
            elif row['Missing Items'] > 0: return ['background-color: #f8d7da'] * len(row)
            elif row['Qty Mismatch'] > 0: return ['background-color: #fff3cd'] * len(row)
            elif row['Quoted Items'] > 0: return ['background-color: #d4edda'] * len(row)
            return [''] * len(row)
        
        st.dataframe(disp_sum.style.apply(color_sum, axis=1), use_container_width=True, hide_index=True)
        
        st.markdown("---")
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📈 Schedule by Code")
            fig1 = px.bar(supplier_summary_df, x='Supplier Code', y='Schedule Line Items', color='Quote Required', text='Schedule Line Items')
            fig1.update_traces(textposition='outside')
            fig1.update_layout(height=400)
            st.plotly_chart(fig1, use_container_width=True)
        
        with col2:
            st.subheader("📊 Quote Coverage")
            cov = supplier_summary_df[supplier_summary_df['Quote Required'] == 'Yes']
            fig2 = go.Figure()
            fig2.add_trace(go.Bar(name='Quoted', x=cov['Supplier Code'], y=cov['Quoted Items'], marker_color='#28a745'))
            fig2.add_trace(go.Bar(name='Missing', x=cov['Supplier Code'], y=cov['Missing Items'], marker_color='#dc3545'))
            fig2.add_trace(go.Bar(name='Mismatch', x=cov['Supplier Code'], y=cov['Qty Mismatch'], marker_color='#ffc107'))
            fig2.update_layout(barmode='stack', height=400)
            st.plotly_chart(fig2, use_container_width=True)

with tabs[4]:
    st.header("💰 Quote Details")
    
    if not st.session_state.quotes:
        st.warning("⚠️ No quotes uploaded yet.")
    else:
        for qname, qdata in st.session_state.quotes.items():
            with st.expander(f"📄 {qname}", expanded=True):
                col1, col2, col3 = st.columns(3)
                col1.markdown(f"**Vendor:** {qdata.get('vendor', 'N/A')}")
                col2.markdown(f"**Date:** {qdata.get('date', 'N/A')}")
                col3.markdown(f"**Total:** ${qdata.get('total', 0):,.2f}")
                
                items_df = pd.DataFrame(qdata.get('items', []))
                if not items_df.empty:
                    disp = items_df.copy()
                    disp['Unit_Price'] = disp['Unit_Price'].apply(lambda x: f"${x:,.2f}")
                    disp['Total'] = disp['Total'].apply(lambda x: f"${x:,.2f}")
                    disp.columns = ['Item #', 'Description', 'Qty', 'Unit Price', 'Total']
                    st.dataframe(disp, use_container_width=True, hide_index=True)

with tabs[5]:
    st.header("📥 Export Report")
    
    if not has_data:
        st.warning("⚠️ Please upload data in the Upload tab first.")
    else:
        st.markdown("""
        **Excel Report includes:** Executive Summary, Supplier Code Summary, Full Analysis, Missing Items, Qty Mismatch, Included Items, Complete Schedule, Quote Details
        """)
        
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Preview")
            included = len(results_df[results_df['Status'] == '✓ Included'])
            missing = len(results_df[results_df['Status'] == '✗ Missing'])
            mismatch = len(results_df[results_df['Status'] == '⚠ Qty Mismatch'])
            st.dataframe(pd.DataFrame({
                "Metric": ["Total Items", "Quoted", "Missing", "Mismatch", "Value"],
                "Value": [len(results_df), included, missing, mismatch, f"${results_df['Total Price'].sum():,.2f}"]
            }), use_container_width=True, hide_index=True)
        
        with col2:
            st.subheader("Supplier Summary")
            st.dataframe(supplier_summary_df[['Supplier Code', 'Schedule Line Items', 'Quoted Items', 'Missing Items']], use_container_width=True, hide_index=True)
        
        st.markdown("---")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            excel = create_excel_report(st.session_state.schedule_df, results_df, st.session_state.quotes, supplier_summary_df, st.session_state.project_name)
            st.download_button("📥 Full Excel Report", excel, f"Quote_Analysis_{datetime.now().strftime('%Y%m%d')}.xlsx",
                              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
        with col2:
            csv = results_df.to_csv(index=False)
            st.download_button("📥 Analysis CSV", csv, "Quote_Analysis.csv", "text/csv", use_container_width=True)
        with col3:
            miss_csv = results_df[results_df['Status'] == '✗ Missing'].to_csv(index=False)
            st.download_button(f"📥 Missing Items ({missing})", miss_csv, "Missing_Items.csv", "text/csv", use_container_width=True)

st.markdown("---")
st.markdown(f"<div style='text-align:center;color:#666'>Kitchen Quote Analyzer v4.0 | {st.session_state.project_name}</div>", unsafe_allow_html=True)
