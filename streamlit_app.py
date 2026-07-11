import streamlit as st
import pandas as pd
import plotly.express as px
import io
import datetime

# --- REPORTLAB IMPORTS FOR THE COMPILATION ENGINE ---
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas

# -------------------- PAGE CONFIG --------------------
st.set_page_config(page_title="Multi-Branch Sales Dashboard", layout="wide")

st.title("📊 Sales Dashboard - Multi-Branch Analysis")
st.markdown("Upload Excel files for **CCK**, **LST**, and **TLT** branches, and download structured executive PDF reports.")

# -------------------- CONSTANTS & CONFIG --------------------
AGENCY_ORDER = ["FGY", "ZB", "APG", "SLA", "JFL", "Others"]
PRODUCT_ORDER = ["FSP", "Pedestal", "Niche", "Others"]
BRANCH_OPTIONS = ["CCK", "LST", "TLT"]

product_colours = {
    "FSP": "#1f77b4", 
    "Pedestal": "#ff7f0e", 
    "Niche": "#2ca02c", 
    "Others": "#d62728"
}

# -------------------- FORMATTING UTILITIES --------------------
def format_currency(value):
    if pd.isna(value): return "$0.00"
    return f"${value:,.2f}"

def format_currency_df(df, numeric_cols):
    df_display = df.copy()
    for col in numeric_cols:
        if col in df_display.columns:
            df_display[col] = df_display[col].apply(format_currency)
    return df_display

def format_chart_label(value):
    if pd.isna(value) or value == 0: return ""
    if abs(value) >= 1e6: return f"${value/1e6:.2f}m"
    elif abs(value) >= 1e3: return f"${value/1e3:.1f}k"
    return f"${value:.2f}"

def format_pct(value):
    if pd.isna(value): return "0.00%"
    return f"{value:,.2f}%"

def sort_by_corporate_hierarchy(df, agency_col="Agency"):
    df[agency_col] = pd.Categorical(df[agency_col], categories=AGENCY_ORDER, ordered=True)
    return df.sort_values(agency_col)

# -------------------- NATIVE PDF REPORT ENGINE --------------------
class NumberedCanvas(canvas.Canvas):
    """Dynamic canvas to handle real-time professional header/footer two-pass pagination."""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_page_decorations(self, page_count):
        self.saveState()
        self.setFont("Helvetica", 8)
        self.setFillColor(colors.HexColor("#555555"))
        
        # Header (Top of page line and timestamp)
        self.setStrokeColor(colors.HexColor("#d3d3d3"))
        self.setLineWidth(0.5)
        self.line(0.75 * inch, 10.5 * inch, 7.75 * inch, 10.5 * inch)
        self.drawString(0.75 * inch, 10.6 * inch, "Corporate Sales Portfolio Performance Matrix")
        
        # Footer
        self.line(0.75 * inch, 0.75 * inch, 7.75 * inch, 0.75 * inch)
        date_str = datetime.date.today().strftime("%d-%b-%Y")
        self.drawString(0.75 * inch, 0.55 * inch, f"Generated on: {date_str} | CONFIDENTIAL")
        self.drawRightString(7.75 * inch, 0.55 * inch, f"Page {self._pageNumber} of {page_count}")
        self.restoreState()

def generate_pdf_report(grand_revenue, total_df, active_branches):
    """Compiles all calculated structural metrics and tables into a clean, presentation-ready PDF document."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=0.5 * inch, rightMargin=0.5 * inch,
        topMargin=0.85 * inch, bottomMargin=0.85 * inch
    )
    
    styles = getSampleStyleSheet()
    
    # Custom corporate font styling configurations
    title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=22, leading=26, textColor=colors.HexColor("#1f77b4"), spaceAfter=6)
    subtitle_style = ParagraphStyle('DocSub', parent=styles['Normal'], fontName='Helvetica-Oblique', fontSize=10, leading=12, textColor=colors.HexColor("#555555"), spaceAfter=15)
    section_style = ParagraphStyle('SecTitle', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=14, leading=18, textColor=colors.HexColor("#2ca02c"), spaceBefore=12, spaceAfter=8)
    table_header_style = ParagraphStyle('TH', fontName='Helvetica-Bold', fontSize=9, leading=11, textColor=colors.white, alignment=1)
    table_cell_style = ParagraphStyle('TC', fontName='Helvetica', fontSize=9, leading=11, textColor=colors.HexColor("#222222"), alignment=1)
    table_cell_bold = ParagraphStyle('TCB', fontName='Helvetica-Bold', fontSize=9, leading=11, textColor=colors.HexColor("#111111"), alignment=1)

    elements = []

    # Title & Metadata Architecture
    elements.append(Paragraph("Consolidated Financial Performance Matrix", title_style))
    elements.append(Paragraph(f"Executive Summary across active branch vectors | Net Corporate Portfolio Revenue: {format_currency(grand_revenue)}", subtitle_style))
    elements.append(Spacer(1, 10))

    # --- SECTION 1: GLOBAL OVERVIEW ---
    elements.append(Paragraph("1. Global Corporate Matrix Summary", section_style))
    
    global_pivot = total_df.pivot_table(index="Agency", columns="Product_Type", values="NETMAINPRODUCT", aggfunc="sum", fill_value=0)
    for p_col in PRODUCT_ORDER:
        if p_col not in global_pivot.columns: global_pivot[p_col] = 0.0
    global_pivot = global_pivot[PRODUCT_ORDER].reindex(AGENCY_ORDER).fillna(0)

    # Build matrix table grid structure
    headers = ["Agency", "FSP", "Pedestal", "Niche", "Others", "Total Sales", "Corp %"]
    table_data = [[Paragraph(h, table_header_style) for h in headers]]
    
    for agency in AGENCY_ORDER:
        fsp = global_pivot.loc[agency, "FSP"]
        ped = global_pivot.loc[agency, "Pedestal"]
        nic = global_pivot.loc[agency, "Niche"]
        oth = global_pivot.loc[agency, "Others"]
        row_tot = fsp + ped + nic + oth
        corp_pct = (row_tot / max(grand_revenue, 1)) * 100
        
        row = [
            Paragraph(agency, table_cell_bold),
            Paragraph(format_currency(fsp), table_cell_style),
            Paragraph(format_currency(ped), table_cell_style),
            Paragraph(format_currency(nic), table_cell_style),
            Paragraph(format_currency(oth), table_cell_style),
            Paragraph(format_currency(row_tot), table_cell_bold),
            Paragraph(format_pct(corp_pct), table_cell_style)
        ]
        table_data.append(row)

    col_widths = [1.1*inch, 1.05*inch, 1.05*inch, 1.05*inch, 1.05*inch, 1.2*inch, 1.0*inch]
    global_table = Table(table_data, colWidths=col_widths, repeatRows=1)
    global_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1f77b4")),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#d3d3d3")),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#f9f9f9")]),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('TOPPADDING', (0,0), (-1,-1), 6),
    ]))
    elements.append(global_table)
    elements.append(Spacer(1, 20))

    # --- SECTION 2: INDIVIDUAL OPERATIONAL BRANCH REGISTRIES ---
    elements.append(Paragraph("2. Operational Breakdown By Branch Entity", section_style))
    
    for b_name, b_df in active_branches.items():
        b_total = b_df["NETMAINPRODUCT"].sum()
        b_pivot = b_df.pivot_table(index="Agency", columns="Product_Type", values="NETMAINPRODUCT", aggfunc="sum", fill_value=0)
        for col in PRODUCT_ORDER:
            if col not in b_pivot.columns: b_pivot[col] = 0.0
        b_pivot = b_pivot[PRODUCT_ORDER].reindex(AGENCY_ORDER).fillna(0)
        
        branch_elements = []
        branch_elements.append(Paragraph(f"Branch Location Matrix: {b_name} (Total: {format_currency(b_total)})", ParagraphStyle('BName', fontName='Helvetica-Bold', fontSize=11, spaceBefore=8, spaceAfter=4, textColor=colors.HexColor("#444444"))))
        
        b_headers = ["Agency", "FSP", "Pedestal", "Niche", "Others", "Branch Vol %", "Global Vol %"]
        b_table_data = [[Paragraph(bh, table_header_style) for bh in b_headers]]
        
        for agency in AGENCY_ORDER:
            fsp = b_pivot.loc[agency, "FSP"]
            ped = b_pivot.loc[agency, "Pedestal"]
            nic = b_pivot.loc[agency, "Niche"]
            oth = b_pivot.loc[agency, "Others"]
            row_tot = fsp + ped + nic + oth
            
            br_vol = (row_tot / max(b_total, 1)) * 100
            gl_vol = (row_tot / max(grand_revenue, 1)) * 100
            
            row = [
                Paragraph(agency, table_cell_bold),
                Paragraph(format_currency(fsp), table_cell_style),
                Paragraph(format_currency(ped), table_cell_style),
                Paragraph(format_currency(nic), table_cell_style),
                Paragraph(format_currency(oth), table_cell_style),
                Paragraph(format_pct(br_vol), table_cell_style),
                Paragraph(format_pct(gl_vol), table_cell_style)
            ]
            b_table_data.append(row)
            
        branch_table = Table(b_table_data, colWidths=col_widths, repeatRows=1)
        branch_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#2ca02c")),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#e2e2e2")),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#fcfcfc")]),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
            ('TOPPADDING', (0,0), (-1,-1), 5),
        ]))
        branch_elements.append(branch_table)
        branch_elements.append(Spacer(1, 12))
        
        # Keep individual branch breakdowns grouped on the same page where possible
        elements.append(KeepTogether(branch_elements))

    # Build the structural document flow via out-canvas logic mapping
    doc.build(elements, canvasmaker=NumberedCanvas)
    return buffer.getvalue()

# -------------------- DATA PROCESSING ENGINES --------------------
def process_branch_file(uploaded_file, branch_name):
    try:
        df_raw = pd.read_excel(uploaded_file, header=None)
        header_row_idx = None
        for idx, row in df_raw.iterrows():
            if row.astype(str).str.contains("FILE_NO").any():
                header_row_idx = idx
                break
        if header_row_idx is None: return None

        df = pd.read_excel(uploaded_file, header=header_row_idx)
        df.columns = df.columns.str.strip()

        required_cols = ["STATUS", "NETMAINPRODUCT", "CBDD_NAME", "BDD_NAME", "PRODUCT_CODE"]
        if any(col not in df.columns for col in required_cols): return None

        df["NETMAINPRODUCT"] = pd.to_numeric(df["NETMAINPRODUCT"], errors="coerce")
        df_confirmed = df[df["STATUS"].str.upper() == "CONFIRM"].copy()
        df_confirmed = df_confirmed[df_confirmed["NETMAINPRODUCT"].notna() & (df_confirmed["NETMAINPRODUCT"] != 0)]

        if df_confirmed.empty: return pd.DataFrame()

        agency_rename = {
            "FU GUI SERVICES": "FGY", "ZENBOX PTE LTD": "ZB",
            "APG ADVISORY PTE. LTD.": "APG", "SINGAPORE LIFESTYLE ASSOCIATES PTE LTD.": "SLA",
            "JF LIFE CONSULTANT PTE LTD": "JFL",
        }

        def get_agency(row):
            cbd = row.get("CBDD_NAME", "")
            if pd.notna(cbd) and str(cbd).strip() != "": raw = str(cbd).strip().upper()
            else:
                bdd = row.get("BDD_NAME", "")
                raw = str(bdd).strip().upper() if pd.notna(bdd) and str(bdd).strip() != "" else "Others"
            for key, value in agency_rename.items():
                if key in raw: return value
            return "Others"

        df_confirmed["Agency"] = df_confirmed.apply(get_agency, axis=1)

        def get_product_type(code):
            if pd.isna(code): return "Others"
            code_str = str(code).strip().upper()
            if code_str == "P": return "FSP"
            elif code_str == "TABLET": return "Pedestal"
            elif code_str == "URN": return "Niche"
            return "Others"

        df_confirmed["Product_Type"] = df_confirmed["PRODUCT_CODE"].apply(get_product_type)
        df_confirmed["Branch"] = branch_name
        
        if "FILE_NO" in df_confirmed.columns:
            df_confirmed = df_confirmed.dropna(subset=["FILE_NO"])
            df_confirmed["FILE_NO"] = df_confirmed["FILE_NO"].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            
        return df_confirmed
    except Exception as e:
        st.error(f"Error processing {branch_name} file: {e}")
        return None

def calculate_cancellations_on_the_fly(editor_state):
    if editor_state is None or editor_state.empty: return 0.0
    valid_records = editor_state.dropna(subset=["Sales_Amount", "Qty"])
    if valid_records.empty: return 0.0
    return (valid_records["Sales_Amount"] * valid_records["Qty"]).sum()

# -------------------- SIDEBAR INTEGRATIONS --------------------
st.sidebar.header("📁 Branch File Uploaders")
cck_file = st.sidebar.file_uploader("Upload CCK Branch Excel", type=["xlsx", "xls"])
lst_file = st.sidebar.file_uploader("Upload LST Branch Excel", type=["xlsx", "xls"])
tlt_file = st.sidebar.file_uploader("Upload TLT Branch Excel", type=["xlsx", "xls"])

st.sidebar.write("---")
st.sidebar.header("🛑 Batch Cancellation Registry")

if "editor_data" not in st.session_state:
    st.session_state.editor_data = pd.DataFrame([
        {"Branch": "LST", "Agency": "FGY", "Product_Type": "Niche", "Sales_Amount": 30000.0, "Qty": 2}
    ])

edited_cxl_df = st.sidebar.data_editor(
    st.session_state.editor_data,
    num_rows="dynamic",
    column_config={
        "Branch": st.column_config.SelectboxColumn("Branch", options=BRANCH_OPTIONS, required=True),
        "Agency": st.column_config.SelectboxColumn("Agency", options=AGENCY_ORDER, required=True),
        "Product_Type": st.column_config.SelectboxColumn("Product Type", options=PRODUCT_ORDER, required=True),
        "Sales_Amount": st.column_config.NumberColumn("Amount ($)", min_value=0.0, format="$%.2f", required=True),
        "Qty": st.column_config.NumberColumn("QTY", min_value=1, step=1, default=1, required=True),
    },
    hide_index=True,
    use_container_width=True,
    key="batch_cxl_editor"
)

pending_total_deduction = calculate_cancellations_on_the_fly(edited_cxl_df)
st.sidebar.metric(label="Active Deductions (On-the-Fly)", value=format_currency(pending_total_deduction))

# -------------------- PROCESS CORE FLOW --------------------
branch_dfs = {}
if cck_file: branch_dfs["CCK"] = process_branch_file(cck_file, "CCK")
if lst_file: branch_dfs["LST"] = process_branch_file(lst_file, "LST")
if tlt_file: branch_dfs["TLT"] = process_branch_file(tlt_file, "TLT")

active_branches = {k: v for k, v in branch_dfs.items() if v is not None and not v.empty}

if not active_branches:
    st.info("👋 Welcome! Please upload branch Excel sheets via the sidebar to populate data views.")
    st.stop()

total_df = pd.concat(active_branches.values(), ignore_index=True)
valid_cxl_df = pd.DataFrame()

if edited_cxl_df is not None and not edited_cxl_df.empty:
    valid_cxl_df = edited_cxl_df.dropna(subset=["Branch", "Agency", "Product_Type", "Sales_Amount", "Qty"])
    for _, item in valid_cxl_df.iterrows():
        row_deduction = item["Sales_Amount"] * item["Qty"]
        if row_deduction <= 0: continue
        
        mask = ((total_df["Branch"] == item["Branch"]) & (total_df["Agency"] == item["Agency"]) & (total_df["Product_Type"] == item["Product_Type"]))
        matching_rows = total_df[mask]
        
        if not matching_rows.empty:
            matching_indices = matching_rows.sort_values(by="NETMAINPRODUCT").index
            deduction_left = row_deduction
            for idx in matching_indices:
                current_val = total_df.at[idx, "NETMAINPRODUCT"]
                if current_val <= deduction_left:
                    deduction_left -= current_val
                    total_df.at[idx, "NETMAINPRODUCT"] = 0.0
                else:
                    total_df.at[idx, "NETMAINPRODUCT"] = current_val - deduction_left
                    deduction_left = 0.0
                    break
    
    total_df = total_df[total_df["NETMAINPRODUCT"] != 0].copy()
    for name in list(active_branches.keys()):
        active_branches[name] = total_df[total_df["Branch"] == name].copy()

grand_corporate_revenue = total_df['NETMAINPRODUCT'].sum()

# -------------------- SIDEBAR DOWNLOAD REGISTRY --------------------
st.sidebar.write("---")
st.sidebar.header("📥 Download Corporate Report")
try:
    pdf_data = generate_pdf_report(grand_corporate_revenue, total_df, active_branches)
    st.sidebar.download_button(
        label="🔴 Download Executive PDF Report",
        data=pdf_data,
        file_name=f"Executive_Sales_Performance_Report.pdf",
        mime="application/pdf",
        use_container_width=True
    )
except Exception as err:
    st.sidebar.error(f"PDF Compiler Standby: {err}")

# -------------------- MAIN DASHBOARD VISUALIZATIONS --------------------
tab_titles = ["Consolidated Overview"] + [f"{name} Branch" for name in active_branches.keys()]
tabs = st.tabs(tab_titles)

with tabs[0]:
    st.header("🏢 Unified Corporate Overview")
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Net Combined Sales", format_currency(grand_corporate_revenue))
    c2.metric("Active Operating Branches", len(active_branches))
    
    st.write("---")
    st.subheader("Overall Performance by Agency Hierarchy")
    if not total_df.empty:
        global_agency_prod = total_df.groupby(["Agency", "Product_Type"])["NETMAINPRODUCT"].sum().reset_index()
        global_agency_prod = sort_by_corporate_hierarchy(global_agency_prod, "Agency")
        fig_gap = px.bar(global_agency_prod, x="Agency", y="NETMAINPRODUCT", color="Product_Type", color_discrete_map=product_colours, barmode="stack", text=global_agency_prod["NETMAINPRODUCT"].apply(format_chart_label))
        st.plotly_chart(fig_gap, use_container_width=True)
    else:
        st.warning("No operating margins left to display.")

for idx, (b_name, b_df) in enumerate(active_branches.items(), start=1):
    with tabs[idx]:
        st.header(f"📍 Operational Analysis: {b_name} Branch")
        if not b_df.empty:
            branch_agency_prod = b_df.groupby(["Agency", "Product_Type"])["NETMAINPRODUCT"].sum().reset_index()
            branch_agency_prod = sort_by_corporate_hierarchy(branch_agency_prod, "Agency")
            fig_br_ap = px.bar(branch_agency_prod, x="Agency", y="NETMAINPRODUCT", color="Product_Type", color_discrete_map=product_colours, barmode="stack", text=branch_agency_prod["NETMAINPRODUCT"].apply(format_chart_label))
            st.plotly_chart(fig_br_ap, use_container_width=True)
