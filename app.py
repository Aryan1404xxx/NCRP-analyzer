import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import plotly.express as px
import plotly.graph_objects as go
import io

st.set_page_config(page_title="NCRP Daily Analyzer", page_icon="🏛️", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp {
    background: linear-gradient(-45deg, #0f0c29, #302b63, #24243e, #1a1a2e);
    background-size: 400% 400%;
    animation: gradientShift 12s ease infinite;
}
@keyframes gradientShift {
    0%   { background-position: 0% 50%; }
    50%  { background-position: 100% 50%; }
    100% { background-position: 0% 50%; }
}
.main-header {
    background: linear-gradient(135deg, #1a3a5c, #2563a8);
    padding: 2rem 2.5rem; border-radius: 16px; margin-bottom: 1.5rem;
}
.main-header h1 { color: white; font-size: 2rem; font-weight: 700; margin: 0; }
.main-header p  { color: rgba(255,255,255,0.75); margin: 0.3rem 0 0; font-size: 0.95rem; }
.metric-card {
    background: rgba(255,255,255,0.07); border: 1px solid rgba(255,255,255,0.15);
    border-radius: 12px; padding: 1rem 1.2rem; text-align: center; backdrop-filter: blur(10px);
}
.metric-num   { font-size: 1.8rem; font-weight: 700; color: white; }
.metric-label { font-size: 0.75rem; color: rgba(255,255,255,0.6); margin-top: 4px; text-transform: uppercase; letter-spacing: 0.05em; }
.section-header {
    font-size: 1rem; font-weight: 600; color: white;
    border-left: 4px solid #2563a8; padding-left: 10px; margin: 1.5rem 0 0.75rem;
}
section[data-testid="stSidebar"] { background: #1a3a5c; }
section[data-testid="stSidebar"] * { color: white !important; }
.stButton > button {
    background: #2563a8; color: white; border: none;
    border-radius: 8px; font-weight: 600; padding: 0.5rem 1.5rem;
}
.stTabs [aria-selected="true"] { color: #60a5fa !important; border-bottom: 3px solid #60a5fa !important; font-weight: 600; }
h1, h2, h3, h4, p, label, .stMarkdown { color: white !important; }
div[data-testid="stMetricValue"] { color: white !important; font-weight: 700 !important; }
div[data-testid="stMetricLabel"] { color: rgba(255,255,255,0.6) !important; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h1>🏛️ NCRP Daily Transaction Analyzer</h1>
    <p>Upload daily NCRP Excel file — instant breakdown of ACKs, accounts, layers, actions, API analysis and more</p>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("## ℹ️ How to Use")
    st.markdown("1️⃣ Upload the daily Excel file")
    st.markdown("2️⃣ Select the sheet with transactions")
    st.markdown("3️⃣ View analysis across all tabs")
    st.markdown("4️⃣ Download full Excel report")
    st.markdown("---")
    st.markdown("**📊 Report Includes**")
    st.markdown("• Daily Summary")
    st.markdown("• Unique ACK Numbers")
    st.markdown("• Unique Account Numbers")
    st.markdown("• Layer-wise Breakdown")
    st.markdown("• Action Breakdown")
    st.markdown("• State-wise Breakdown")
    st.markdown("• Bank Breakdown")
    st.markdown("• Officer Breakdown")
    st.markdown("• API Deep Analysis")
    st.markdown("---")
    st.markdown("**Developed by:** Aryan Sinha")

@st.cache_data
def load_data(file_bytes, sheet_name):
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name)
    df.columns = [str(c).strip() for c in df.columns]
    df['Transaction Amount'] = pd.to_numeric(df.get('Transaction Amount', 0), errors='coerce').fillna(0)
    df['Disputed Amount']    = pd.to_numeric(df.get('Disputed Amount', 0),    errors='coerce').fillna(0)
    df['Layers']             = pd.to_numeric(df.get('Layers', 0),             errors='coerce').fillna(0)
    if 'Acknowledgement No.' in df.columns:
        df['Acknowledgement No.'] = df['Acknowledgement No.'].astype(str).str.strip()
    if 'Account Number' in df.columns:
        df['Account Number'] = df['Account Number'].astype(str).str.strip()
    if 'Action' in df.columns:
        df['Action'] = df['Action'].astype(str).str.strip()
    return df

def layer_cat(x):
    if x == 0: return 'Layer 0'
    elif x <= 3: return 'Layer 1-3'
    elif x <= 5: return 'Layer 4-5'
    elif x <= 10: return 'Layer 6-10'
    else: return 'Layer 10+'

def build_excel(df, date_label, api_df, cat_df, api_layer, api_state, api_bank, api_officer):
    wb = openpyxl.Workbook()
    NAVY   = PatternFill("solid", fgColor="1a3a5c")
    BLUE   = PatternFill("solid", fgColor="2563a8")
    GREEN  = PatternFill("solid", fgColor="059669")
    RED    = PatternFill("solid", fgColor="DC2626")
    ORANGE = PatternFill("solid", fgColor="D97706")
    PURPLE = PatternFill("solid", fgColor="7C3AED")
    ALT1   = PatternFill("solid", fgColor="EFF6FF")
    WHITE  = PatternFill("solid", fgColor="FFFFFF")
    hfont  = Font(name='Arial', bold=True, color='FFFFFF', size=10)
    dfont  = Font(name='Arial', size=10)
    tfont  = Font(name='Arial', bold=True, color='FFFFFF', size=13)
    bfont  = Font(name='Arial', bold=True, size=11)
    border = Border(
        left=Side(style='thin', color='E5E7EB'), right=Side(style='thin', color='E5E7EB'),
        top=Side(style='thin', color='E5E7EB'),  bottom=Side(style='thin', color='E5E7EB')
    )
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left   = Alignment(horizontal='left',   vertical='center', wrap_text=True)

    def title_row(ws, text, fill, ncols, row=1):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
        c = ws.cell(row=row, column=1, value=text)
        c.fill=fill; c.font=tfont; c.alignment=center; c.border=border
        ws.row_dimensions[row].height = 28

    def header_row(ws, headers, fill, row=2):
        for col, h in enumerate(headers, 1):
            c = ws.cell(row=row, column=col, value=h)
            c.fill=fill; c.font=hfont; c.alignment=center; c.border=border
        ws.row_dimensions[row].height = 20

    def data_row(ws, values, row, alt=True):
        fill = ALT1 if alt else WHITE
        for col, v in enumerate(values, 1):
            c = ws.cell(row=row, column=col, value=v)
            c.fill=fill; c.font=dfont; c.alignment=left; c.border=border
        ws.row_dimensions[row].height = 18

    def breakdown_sheet(ws, title, data, col_headers, fill, col_widths):
        title_row(ws, title, fill, len(col_headers))
        header_row(ws, col_headers, fill)
        for i, (_, row) in enumerate(data.iterrows(), 3):
            vals = [row[c] for c in data.columns]
            data_row(ws, vals, i, i%2==0)
        for col, w in zip([get_column_letter(i+1) for i in range(len(col_widths))], col_widths):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = 'A3'

    total_txns = len(df)
    unique_ack = df['Acknowledgement No.'].nunique() if 'Acknowledgement No.' in df.columns else 0
    unique_acc = df['Account Number'].nunique() if 'Account Number' in df.columns else 0
    api_count  = (df['Action'] == 'Money transfer TO').sum() if 'Action' in df.columns else 0
    ccc_count  = total_txns - api_count
    api_pct    = round(api_count/total_txns*100, 2) if total_txns else 0
    layer0     = (df['Layers']==0).sum()
    layer1_3   = df['Layers'].between(1,3).sum()
    layer4_5   = df['Layers'].between(4,5).sum()
    layer6_10  = df['Layers'].between(6,10).sum()
    layer10p   = (df['Layers']>10).sum()
    dis_lt500  = (df['Disputed Amount']<500).sum()
    txn_lt500  = (df['Transaction Amount']<500).sum()
    txn_20kp   = (df['Transaction Amount']>=20000).sum()

    # Sheet 1 - Summary
    ws1 = wb.active; ws1.title = "Daily Summary"
    title_row(ws1, f'NCRP DAILY ANALYSIS — {date_label} | Total: {total_txns:,}', NAVY, 4)
    ws1.merge_cells('A3:B3')
    c=ws1.cell(row=3,column=1,value="METRIC"); c.fill=NAVY; c.font=hfont; c.alignment=center; c.border=border
    ws1.merge_cells('C3:D3')
    c=ws1.cell(row=3,column=3,value="VALUE"); c.fill=NAVY; c.font=hfont; c.alignment=center; c.border=border
    metrics = [
        ("Total Transactions", total_txns, BLUE),
        ("Unique ACK Numbers", unique_ack, GREEN),
        ("Unique Account Numbers", unique_acc, PURPLE),
        ("API (Money Transfer TO)", api_count, ORANGE),
        ("CCC", ccc_count, BLUE),
        ("API %", f"{api_pct}%", GREEN),
        ("Layer 0", layer0, NAVY),
        ("Layer 1-3", layer1_3, BLUE),
        ("Layer 4-5", layer4_5, ORANGE),
        ("Layer 6-10", layer6_10, RED),
        ("Layer 10+", layer10p, RED),
        ("Disputed Amount < ₹500", dis_lt500, GREEN),
        ("Transaction Amount < ₹500", txn_lt500, GREEN),
        ("Transaction Amount ≥ ₹20,000", txn_20kp, ORANGE),
    ]
    for i, (label, val, fill) in enumerate(metrics, 4):
        ws1.merge_cells(start_row=i,start_column=1,end_row=i,end_column=2)
        c1=ws1.cell(row=i,column=1,value=label)
        c1.fill=ALT1 if i%2==0 else WHITE; c1.font=bfont; c1.alignment=left; c1.border=border
        ws1.merge_cells(start_row=i,start_column=3,end_row=i,end_column=4)
        c2=ws1.cell(row=i,column=3,value=val)
        c2.fill=fill; c2.font=Font(name='Arial',bold=True,color='FFFFFF',size=12)
        c2.alignment=center; c2.border=border
        ws1.row_dimensions[i].height=22
    for col,w in zip(['A','B','C','D'],[30,10,20,10]):
        ws1.column_dimensions[col].width=w

    # Sheet 2 - Action
    if 'Action' in df.columns:
        ac = df['Action'].value_counts().reset_index(); ac.columns=['Action','Count']
        ac['% of Total'] = (ac['Count']/total_txns*100).round(2).astype(str)+'%'
        ws2 = wb.create_sheet("Action Breakdown")
        breakdown_sheet(ws2,'ACTION BREAKDOWN',ac,['Action','Count','% of Total'],BLUE,[40,12,14])

    # Sheet 3 - Layer
    ld = pd.DataFrame({
        'Layer Category':['Layer 0','Layer 1-3','Layer 4-5','Layer 6-10','Layer 10+'],
        'Count':[layer0,layer1_3,layer4_5,layer6_10,layer10p]
    })
    ld['% of Total'] = (ld['Count']/total_txns*100).round(2).astype(str)+'%'
    ws3 = wb.create_sheet("Layer Breakdown")
    breakdown_sheet(ws3,'LAYER-WISE BREAKDOWN',ld,['Layer Category','Count','% of Total'],PURPLE,[20,12,14])

    # Sheet 4 - State
    if 'State' in df.columns:
        sc = df['State'].value_counts().reset_index(); sc.columns=['State','Count']
        sc['% of Total'] = (sc['Count']/total_txns*100).round(2).astype(str)+'%'
        ws4 = wb.create_sheet("State Breakdown")
        breakdown_sheet(ws4,'STATE-WISE BREAKDOWN',sc,['State','Count','% of Total'],GREEN,[35,12,14])

    # Sheet 5 - Bank
    if 'Money transfer TO Bank' in df.columns:
        bc = df['Money transfer TO Bank'].value_counts().reset_index(); bc.columns=['Bank','Count']
        bc['% of Total'] = (bc['Count']/total_txns*100).round(2).astype(str)+'%'
        ws5 = wb.create_sheet("Bank Breakdown")
        breakdown_sheet(ws5,'BANK BREAKDOWN',bc,['Bank','Count','% of Total'],ORANGE,[60,12,14])

    # Sheet 6 - Officer
    if 'Action Taken by' in df.columns:
        oc = df['Action Taken by'].value_counts().reset_index(); oc.columns=['Officer','Count']
        oc['% of Total'] = (oc['Count']/total_txns*100).round(2).astype(str)+'%'
        ws6 = wb.create_sheet("Officer Breakdown")
        breakdown_sheet(ws6,'OFFICER BREAKDOWN',oc,['Officer','Count','% of Total'],NAVY,[35,12,14])

    # Sheet 7 - Unique ACK
    if 'Acknowledgement No.' in df.columns:
        ack = df.groupby('Acknowledgement No.').agg(
            Transactions=('Acknowledgement No.','count'),
            Total_Amount=('Transaction Amount','sum'),
            Disputed_Amount=('Disputed Amount','sum'),
            Layers=('Layers','first'),
            State=('State','first') if 'State' in df.columns else ('Layers','first'),
            Action=('Action','first') if 'Action' in df.columns else ('Layers','first'),
        ).reset_index()
        ws7 = wb.create_sheet("Unique ACK List")
        title_row(ws7,f'UNIQUE ACK NUMBERS — {unique_ack:,}',RED,6)
        header_row(ws7,['ACK No.','Transactions','Total Amount','Disputed Amount','State','Action'],RED)
        for i,(_, row) in enumerate(ack.iterrows(),3):
            data_row(ws7,[row['Acknowledgement No.'],row['Transactions'],row['Total_Amount'],
                          row['Disputed_Amount'],row.get('State',''),row.get('Action','')],i,i%2==0)
        for col,w in zip(['A','B','C','D','E','F'],[18,14,16,16,25,30]):
            ws7.column_dimensions[col].width=w
        ws7.freeze_panes='A3'

    # Sheet 8 - Unique Account
    if 'Account Number' in df.columns:
        acc = df.groupby('Account Number').agg(
            Transactions=('Account Number','count'),
            Total_Amount=('Transaction Amount','sum'),
            Disputed_Amount=('Disputed Amount','sum'),
            Unique_ACKs=('Acknowledgement No.','nunique') if 'Acknowledgement No.' in df.columns else ('Account Number','count'),
            State=('State','first') if 'State' in df.columns else ('Account Number','count'),
            Action=('Action','first') if 'Action' in df.columns else ('Account Number','count'),
        ).reset_index()
        ws8 = wb.create_sheet("Unique Account List")
        title_row(ws8,f'UNIQUE ACCOUNT NUMBERS — {unique_acc:,}',PURPLE,7)
        header_row(ws8,['Account Number','Transactions','Total Amount','Disputed Amount','Unique ACKs','State','Action'],PURPLE)
        for i,(_, row) in enumerate(acc.iterrows(),3):
            data_row(ws8,[row['Account Number'],row['Transactions'],row['Total_Amount'],
                          row['Disputed_Amount'],row.get('Unique_ACKs',''),
                          row.get('State',''),row.get('Action','')],i,i%2==0)
        for col,w in zip(['A','B','C','D','E','F','G'],[22,14,16,16,12,25,30]):
            ws8.column_dimensions[col].width=w
        ws8.freeze_panes='A3'

    # Sheet 9 - API Summary
    ws9 = wb.create_sheet("API Analysis")
    title_row(ws9,f'API ANALYSIS — {len(api_df):,} API transactions',BLUE,3)
    header_row(ws9,['Category','Count','% of Total'],BLUE)
    for i,(_, row) in enumerate(cat_df.iterrows(),3):
        data_row(ws9,[row.get('Action', row.get('Category','')),row['Count'],row.get('% of API Total', row.get('% of Total',''))],i,i%2==0)
    for col,w in zip(['A','B','C'],[30,12,14]):
        ws9.column_dimensions[col].width=w

    ws10 = wb.create_sheet("API by Layer")
    breakdown_sheet(ws10,'API BY LAYER',api_layer,['Layer','API Count','% of API'],PURPLE,[20,12,14])

    if api_state is not None:
        ws11 = wb.create_sheet("API by State")
        breakdown_sheet(ws11,'API BY STATE',api_state,['State','API Count','% of API'],GREEN,[35,12,14])

    if api_bank is not None:
        ws12 = wb.create_sheet("API by Bank")
        breakdown_sheet(ws12,'API BY BANK',api_bank,['Bank','API Count','% of API'],ORANGE,[60,12,14])

    if api_officer is not None:
        ws13 = wb.create_sheet("API by Officer")
        breakdown_sheet(ws13,'API BY OFFICER',api_officer,['Officer','API Count','% of API'],NAVY,[35,12,14])

    # Sheet 14 - Raw API
    ws14 = wb.create_sheet("Raw API Transactions")
    api_raw = api_df.copy()
    title_row(ws14,f'RAW API TRANSACTIONS — {len(api_raw):,}',BLUE,len(api_raw.columns))
    for col,h in enumerate(api_raw.columns,1):
        c=ws14.cell(row=2,column=col,value=h)
        c.fill=BLUE; c.font=hfont; c.alignment=center; c.border=border
    for i,(_, row) in enumerate(api_raw.iterrows(),3):
        fill=ALT1 if i%2==0 else WHITE
        for col,val in enumerate(row,1):
            c=ws14.cell(row=i,column=col,value=val)
            c.fill=fill; c.font=dfont; c.alignment=left; c.border=border
        ws14.row_dimensions[i].height=16
    for col in range(1,len(api_raw.columns)+1):
        ws14.column_dimensions[get_column_letter(col)].width=18
    ws14.freeze_panes='A3'

    # Sheet 15 - Raw Data
    ws15 = wb.create_sheet("Raw Data")
    title_row(ws15,'RAW TRANSACTION DATA',NAVY,len(df.columns))
    for col,h in enumerate(df.columns,1):
        c=ws15.cell(row=2,column=col,value=h)
        c.fill=NAVY; c.font=hfont; c.alignment=center; c.border=border
    for i,(_, row) in enumerate(df.iterrows(),3):
        fill=ALT1 if i%2==0 else WHITE
        for col,val in enumerate(row,1):
            c=ws15.cell(row=i,column=col,value=val)
            c.fill=fill; c.font=dfont; c.alignment=left; c.border=border
        ws15.row_dimensions[i].height=16
    for col in range(1,len(df.columns)+1):
        ws15.column_dimensions[get_column_letter(col)].width=20
    ws15.freeze_panes='A3'

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf

# ---- MAIN UI ----
uploaded = st.file_uploader("📂 Upload daily NCRP Excel file (.xlsx)", type=["xlsx"])

if uploaded is None:
    st.info("👆 Upload the daily NCRP Excel file to get started")
    sample = pd.DataFrame({
        'Acknowledgement No.': ['20107260000485','20107260000485'],
        'Transaction Amount':  [98000, 90000],
        'Disputed Amount':     [0, 0],
        'Layers':              [0, 0],
        'Action':              ['Money transfer TO','Transaction put on hold'],
        'State':               ['Delhi','Mumbai'],
        'Account Number':      ['ptminj-1@oksbi','ptminj-1@oksbi'],
    })
    st.dataframe(sample, use_container_width=True)

else:
    file_bytes = uploaded.read()
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets = xl.sheet_names
    selected_sheet = st.selectbox("📄 Select sheet", sheets,
        index=sheets.index('ViewTransactions') if 'ViewTransactions' in sheets else 0)
    date_label = st.text_input("📅 Date label for report", value="Daily Analysis")

    with st.spinner("📊 Analyzing..."):
        df = load_data(file_bytes, selected_sheet)

    if df.empty:
        st.error("❌ No data found")
    else:
        total_txns = len(df)
        unique_ack = df['Acknowledgement No.'].nunique() if 'Acknowledgement No.' in df.columns else 0
        unique_acc = df['Account Number'].nunique() if 'Account Number' in df.columns else 0
        api_count  = (df['Action'] == 'Money transfer TO').sum() if 'Action' in df.columns else 0
        ccc_count  = total_txns - api_count
        api_pct    = round(api_count/total_txns*100, 2) if total_txns else 0
        layer0     = (df['Layers']==0).sum()
        layer1_3   = df['Layers'].between(1,3).sum()
        layer4_5   = df['Layers'].between(4,5).sum()
        layer6_10  = df['Layers'].between(6,10).sum()
        layer10p   = (df['Layers']>10).sum()

        # API data
        df['Layer_Cat'] = df['Layers'].apply(layer_cat)
        api_mask = df['Action Taken by'].astype(str).str.upper().str.contains('API', na=False) if 'Action Taken by' in df.columns else pd.Series([False]*len(df))
        api_df = df[api_mask].copy()
        api_n  = len(api_df)
        ccc_df = df[~api_mask].copy()

        cat_df = pd.DataFrame()
        if 'Action' in df.columns and not api_df.empty:
            cat_df = api_df['Action'].value_counts().reset_index()
            cat_df.columns = ['Action', 'Count']
            cat_df['% of API Total'] = (cat_df['Count']/api_n*100).round(2).astype(str)+'%'

        if not api_df.empty:
            api_layer = api_df['Layer_Cat'].value_counts().reset_index()
            api_layer.columns = ['Layer','API Count']
            order = ['Layer 0','Layer 1-3','Layer 4-5','Layer 6-10','Layer 10+']
            api_layer['Layer'] = pd.Categorical(api_layer['Layer'], categories=order, ordered=True)
            api_layer = api_layer.sort_values('Layer')
            api_layer['% of API'] = (api_layer['API Count']/api_n*100).round(2).astype(str)+'%'

            api_state = None
            if 'State' in api_df.columns:
                api_state = api_df['State'].value_counts().reset_index()
                api_state.columns = ['State','API Count']
                api_state['% of API'] = (api_state['API Count']/api_n*100).round(2).astype(str)+'%'

            api_bank = None
            if 'Money transfer TO Bank' in api_df.columns:
                api_bank = api_df['Money transfer TO Bank'].value_counts().reset_index()
                api_bank.columns = ['Bank','API Count']
                api_bank['% of API'] = (api_bank['API Count']/api_n*100).round(2).astype(str)+'%'

            api_officer = None
            if 'Action Taken by' in api_df.columns:
                api_officer = api_df['Action Taken by'].value_counts().reset_index()
                api_officer.columns = ['Officer','API Count']
                api_officer['% of API'] = (api_officer['API Count']/api_n*100).round(2).astype(str)+'%'
        else:
            api_layer = pd.DataFrame()
            api_state = api_bank = api_officer = None

        # Metrics row
        c1,c2,c3,c4,c5,c6 = st.columns(6)
        with c1: st.markdown(f'<div class="metric-card"><div class="metric-num">{total_txns:,}</div><div class="metric-label">Total Transactions</div></div>', unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="metric-card"><div class="metric-num">{unique_ack:,}</div><div class="metric-label">Unique ACKs</div></div>', unsafe_allow_html=True)
        with c3: st.markdown(f'<div class="metric-card"><div class="metric-num">{unique_acc:,}</div><div class="metric-label">Unique Accounts</div></div>', unsafe_allow_html=True)
        with c4: st.markdown(f'<div class="metric-card"><div class="metric-num" style="color:#60a5fa">{api_count:,}</div><div class="metric-label">API</div></div>', unsafe_allow_html=True)
        with c5: st.markdown(f'<div class="metric-card"><div class="metric-num">{ccc_count:,}</div><div class="metric-label">CCC</div></div>', unsafe_allow_html=True)
        with c6: st.markdown(f'<div class="metric-card"><div class="metric-num" style="color:#34d399">{api_pct}%</div><div class="metric-label">API %</div></div>', unsafe_allow_html=True)

        st.markdown("---")

        tab1,tab2,tab3,tab4,tab5,tab6,tab7,tab8 = st.tabs([
            "📊 Summary","🔢 ACK & Accounts","🧅 Layers","⚡ Actions","🗺️ States","🏦 Banks","👮 Officers","🎯 API Analysis"
        ])

        with tab1:
            st.markdown('<div class="section-header">Layer Distribution</div>', unsafe_allow_html=True)
            layer_df = pd.DataFrame({
                'Layer': ['Layer 0','Layer 1-3','Layer 4-5','Layer 6-10','Layer 10+'],
                'Count': [layer0, layer1_3, layer4_5, layer6_10, layer10p]
            })
            col_a, col_b = st.columns(2)
            with col_a:
                fig = px.bar(layer_df, x='Layer', y='Count', color='Layer',
                    color_discrete_sequence=['#2563a8','#059669','#D97706','#DC2626','#7C3AED'],
                    title='Transactions by Layer')
                fig.update_layout(showlegend=False, height=300, paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                st.plotly_chart(fig, use_container_width=True)
            with col_b:
                fig2 = px.pie(layer_df, values='Count', names='Layer', hole=0.4,
                    title='Layer Distribution %',
                    color_discrete_sequence=['#2563a8','#059669','#D97706','#DC2626','#7C3AED'])
                fig2.update_layout(height=300, paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                st.plotly_chart(fig2, use_container_width=True)
            st.markdown('<div class="section-header">Amount Analysis</div>', unsafe_allow_html=True)
            a1,a2,a3 = st.columns(3)
            a1.metric("Disputed Amount < ₹500", f"{(df['Disputed Amount']<500).sum():,}")
            a2.metric("Transaction Amount < ₹500", f"{(df['Transaction Amount']<500).sum():,}")
            a3.metric("Transaction Amount ≥ ₹20,000", f"{(df['Transaction Amount']>=20000).sum():,}")

        with tab2:
            col_a, col_b = st.columns(2)
            with col_a:
                st.markdown('<div class="section-header">Unique ACK Numbers</div>', unsafe_allow_html=True)
                if 'Acknowledgement No.' in df.columns:
                    ack_df = df.groupby('Acknowledgement No.').agg(
                        Transactions=('Acknowledgement No.','count'),
                        Total_Amount=('Transaction Amount','sum'),
                        Disputed_Amount=('Disputed Amount','sum'),
                        Layers=('Layers','first'),
                    ).reset_index()
                    st.metric("Total Unique ACKs", f"{len(ack_df):,}")
                    st.dataframe(ack_df, use_container_width=True, height=350)
            with col_b:
                st.markdown('<div class="section-header">Unique Account Numbers</div>', unsafe_allow_html=True)
                if 'Account Number' in df.columns:
                    acc_df = df.groupby('Account Number').agg(
                        Transactions=('Account Number','count'),
                        Total_Amount=('Transaction Amount','sum'),
                        Disputed_Amount=('Disputed Amount','sum'),
                        Unique_ACKs=('Acknowledgement No.','nunique') if 'Acknowledgement No.' in df.columns else ('Account Number','count'),
                    ).reset_index()
                    st.metric("Total Unique Accounts", f"{len(acc_df):,}")
                    st.dataframe(acc_df, use_container_width=True, height=350)

        with tab3:
            st.markdown('<div class="section-header">Layer-wise Breakdown</div>', unsafe_allow_html=True)
            layer_detail = pd.DataFrame({
                'Layer Category': ['Layer 0','Layer 1-3','Layer 4-5','Layer 6-10','Layer 10+'],
                'Count':          [layer0, layer1_3, layer4_5, layer6_10, layer10p],
                '% of Total':     [f"{round(x/total_txns*100,2)}%" for x in [layer0,layer1_3,layer4_5,layer6_10,layer10p]]
            })
            st.dataframe(layer_detail, use_container_width=True)
            fig3 = px.bar(layer_detail, x='Layer Category', y='Count', text='% of Total', color='Layer Category',
                color_discrete_sequence=['#2563a8','#059669','#D97706','#DC2626','#7C3AED'])
            fig3.update_layout(showlegend=False, height=300, paper_bgcolor='rgba(0,0,0,0)', font_color='white')
            st.plotly_chart(fig3, use_container_width=True)

            if not api_df.empty:
                st.markdown('<div class="section-header">API (Money Transfer TO) per Layer</div>', unsafe_allow_html=True)
                col_a, col_b = st.columns(2)
                with col_a:
                    st.dataframe(api_layer, use_container_width=True)
                with col_b:
                    fig_al = px.bar(api_layer, x='Layer', y='API Count', text='% of API',
                        color='Layer', title='API Transactions by Layer',
                        color_discrete_sequence=['#2563a8','#059669','#D97706','#DC2626','#7C3AED'])
                    fig_al.update_layout(showlegend=False, height=300, paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                    st.plotly_chart(fig_al, use_container_width=True)

        with tab4:
            st.markdown('<div class="section-header">Action Breakdown</div>', unsafe_allow_html=True)
            if 'Action' in df.columns:
                ac = df['Action'].value_counts().reset_index(); ac.columns=['Action','Count']
                ac['% of Total'] = (ac['Count']/total_txns*100).round(2).astype(str)+'%'
                col_a, col_b = st.columns(2)
                with col_a:
                    st.dataframe(ac, use_container_width=True, height=350)
                with col_b:
                    fig4 = px.pie(ac, values='Count', names='Action', hole=0.4, title='Action Distribution')
                    fig4.update_layout(height=350, paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                    st.plotly_chart(fig4, use_container_width=True)

        with tab5:
            st.markdown('<div class="section-header">State-wise Breakdown</div>', unsafe_allow_html=True)
            if 'State' in df.columns:
                sc = df['State'].value_counts().reset_index(); sc.columns=['State','Count']
                sc['% of Total'] = (sc['Count']/total_txns*100).round(2).astype(str)+'%'
                col_a, col_b = st.columns(2)
                with col_a:
                    st.dataframe(sc, use_container_width=True, height=400)
                with col_b:
                    fig5 = px.bar(sc.head(15), x='Count', y='State', orientation='h',
                        title='Top 15 States', color='Count',
                        color_continuous_scale=['#EFF6FF','#2563a8'])
                    fig5.update_layout(height=400, paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                    st.plotly_chart(fig5, use_container_width=True)

        with tab6:
            st.markdown('<div class="section-header">Bank Breakdown</div>', unsafe_allow_html=True)
            if 'Money transfer TO Bank' in df.columns:
                bc = df['Money transfer TO Bank'].value_counts().reset_index(); bc.columns=['Bank','Count']
                bc['% of Total'] = (bc['Count']/total_txns*100).round(2).astype(str)+'%'
                col_a, col_b = st.columns(2)
                with col_a:
                    st.dataframe(bc, use_container_width=True, height=400)
                with col_b:
                    fig6 = px.bar(bc.head(10), x='Count', y='Bank', orientation='h',
                        title='Top 10 Banks', color='Count',
                        color_continuous_scale=['#FFF7ED','#D97706'])
                    fig6.update_layout(height=400, paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                    st.plotly_chart(fig6, use_container_width=True)

        with tab7:
            st.markdown('<div class="section-header">Officer-wise Action Breakdown</div>', unsafe_allow_html=True)
            if 'Action Taken by' in df.columns:
                oc = df['Action Taken by'].value_counts().reset_index(); oc.columns=['Officer','Count']
                oc['% of Total'] = (oc['Count']/total_txns*100).round(2).astype(str)+'%'
                col_a, col_b = st.columns(2)
                with col_a:
                    st.dataframe(oc, use_container_width=True, height=400)
                with col_b:
                    fig7 = px.bar(oc.head(15), x='Count', y='Officer', orientation='h',
                        title='Officer Performance', color='Count',
                        color_continuous_scale=['#F5F3FF','#7C3AED'])
                    fig7.update_layout(height=400, paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                    st.plotly_chart(fig7, use_container_width=True)

        with tab8:
            st.markdown('<div class="section-header">API vs Put on Hold vs Others</div>', unsafe_allow_html=True)
            c1, c2, c3 = st.columns(3)
            c1.metric("🔵 Total API Transactions", f"{api_n:,}", f"{round(api_n/total_txns*100,2)}% of total")
            c2.metric("⚪ Total CCC Transactions", f"{len(ccc_df):,}", f"{round(len(ccc_df)/total_txns*100,2)}% of total")
            c3.metric("📊 Total Transactions", f"{total_txns:,}")

            if not cat_df.empty:
                st.markdown('<div class="section-header">API Action Breakdown</div>', unsafe_allow_html=True)
                col_a, col_b = st.columns(2)
                with col_a:
                    st.dataframe(cat_df, use_container_width=True)
                with col_b:
                    fig_cat = px.pie(cat_df, values='Count', names='Action', hole=0.4,
                        title='API Actions Breakdown',
                        color_discrete_sequence=['#2563a8','#DC2626','#D97706','#059669','#7C3AED','#6B7280'])
                    fig_cat.update_layout(height=320, paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                    st.plotly_chart(fig_cat, use_container_width=True)

            if not api_df.empty:
                st.markdown('<div class="section-header">API by Layer</div>', unsafe_allow_html=True)
                col_a, col_b = st.columns(2)
                with col_a:
                    st.dataframe(api_layer, use_container_width=True)
                with col_b:
                    fig_al2 = px.bar(api_layer, x='Layer', y='API Count', text='% of API',
                        color='Layer', title='API by Layer',
                        color_discrete_sequence=['#2563a8','#059669','#D97706','#DC2626','#7C3AED'])
                    fig_al2.update_layout(showlegend=False, height=300, paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                    st.plotly_chart(fig_al2, use_container_width=True)

                if api_state is not None:
                    st.markdown('<div class="section-header">API by State</div>', unsafe_allow_html=True)
                    col_a, col_b = st.columns(2)
                    with col_a:
                        st.dataframe(api_state, use_container_width=True, height=350)
                    with col_b:
                        fig_as = px.bar(api_state.head(12), x='API Count', y='State', orientation='h',
                            title='Top States — API', color='API Count',
                            color_continuous_scale=['#EFF6FF','#2563a8'])
                        fig_as.update_layout(height=380, paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                        st.plotly_chart(fig_as, use_container_width=True)

                if api_bank is not None:
                    st.markdown('<div class="section-header">API by Bank</div>', unsafe_allow_html=True)
                    col_a, col_b = st.columns(2)
                    with col_a:
                        st.dataframe(api_bank, use_container_width=True, height=350)
                    with col_b:
                        fig_ab = px.bar(api_bank.head(10), x='API Count', y='Bank', orientation='h',
                            title='Top Banks — API', color='API Count',
                            color_continuous_scale=['#FFF7ED','#D97706'])
                        fig_ab.update_layout(height=380, paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                        st.plotly_chart(fig_ab, use_container_width=True)

                if api_officer is not None:
                    st.markdown('<div class="section-header">API by Officer</div>', unsafe_allow_html=True)
                    col_a, col_b = st.columns(2)
                    with col_a:
                        st.dataframe(api_officer, use_container_width=True, height=350)
                    with col_b:
                        fig_ao = px.bar(api_officer.head(12), x='API Count', y='Officer', orientation='h',
                            title='Top Officers — API', color='API Count',
                            color_continuous_scale=['#F5F3FF','#7C3AED'])
                        fig_ao.update_layout(height=380, paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                        st.plotly_chart(fig_ao, use_container_width=True)

                st.markdown("---")
                api_excel = io.BytesIO()
                with pd.ExcelWriter(api_excel, engine='openpyxl') as writer:
                    cat_df.to_excel(writer, sheet_name='API vs Hold vs Others', index=False)
                    api_layer.to_excel(writer, sheet_name='API by Layer', index=False)
                    if api_state is not None:
                        api_state.to_excel(writer, sheet_name='API by State', index=False)
                    if api_bank is not None:
                        api_bank.to_excel(writer, sheet_name='API by Bank', index=False)
                    if api_officer is not None:
                        api_officer.to_excel(writer, sheet_name='API by Officer', index=False)
                    api_df.to_excel(writer, sheet_name='Raw API Transactions', index=False)
                api_excel.seek(0)
                st.download_button(
                    "📥 Download API Analysis Report",
                    api_excel,
                    "API_Analysis_Report.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

        st.markdown("---")
        with st.spinner("📥 Preparing full Excel report..."):
            excel_buf = build_excel(df, date_label, api_df, cat_df, api_layer,
                                    api_state, api_bank, api_officer)
        st.download_button(
            "📥 Download Full Excel Report",
            excel_buf,
            f"NCRP_Analysis_{date_label.replace(' ','_')}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
