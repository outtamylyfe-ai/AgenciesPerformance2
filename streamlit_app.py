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
st.markdown("Upload Excel files for **CCK**, **LST**, and **TLT** branches, and download structured executive PDF reports with monthly insights.")

# -------------------- CONSTANTS & CONFIG --------------------
AGENCY_ORDER = ["FGY", "ZB", "APG", "SLA", "JFL", "Others"]
PRODUCT_ORDER = ["FSP", "Pedestal", "Niche", "Others"]
BRANCH_OPTIONS = ["CCK", "LST", "TLT"]
MONTH_ORDER = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

product_colours = {
    "FSP": "#1f77b4", 
    "Pedestal": "#ff7f0e", 
    "Niche": "#2ca02c", 
    "Others": "#d62728"
}

# -------------------- FORMATTING UTILITIES --------------------
def format_currency(value):
    if pd.isna(value) or value is None: return "$0.00"
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

def sort_by_month(df, month_col="Month"):
    df[month_col] = pd.Categorical(df[month_col], categories=MONTH_ORDER, ordered=True)
    return df.sort_values(month_col)

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
        
        # Header Bound Line
        self.setStrokeColor(colors.HexColor("#d3d3d3"))
        self.setLineWidth(0.5)
        self.line(0.5 * inch, 10.5 * inch, 8.0 * inch, 10.5 * inch)
        self.drawString(0.5 * inch, 10.6 * inch, "Executive Performance Overview")
        
        # Footer
        self.line(0.5 * inch, 0.75 * inch, 8.0 * inch, 0.75 * inch)
        date_str = datetime.date.today().strftime("%d-%b-%Y")
        self.drawString(0.5 * inch, 0.55 * inch, f"Generated on: {date_str} | EXECUTIVE DATA PREVIEW")
        self.drawRightString(8.0 * inch, 0.55 * inch, f"Page {self._pageNumber} of {page_count}")
        self.restoreState()

def generate_pdf_report(grand_revenue, total_df, active_branches, valid_cxl_df, total_deducted):
    """Compiles all structured matrix arrays and data tables into an elegant PDF layout without charts."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=0.5 * inch, rightMargin=0.5 * inch,
        topMargin=0.85 * inch, bottomMargin=0.85 * inch
    )
    
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=20, leading=24, textColor=colors.HexColor("#1f77b4"), spaceAfter=4)
    subtitle_style = ParagraphStyle('DocSub', parent=styles['Normal'], fontName='Helvetica-Oblique', fontSize=9, leading=11, textColor=colors.HexColor("#555555"), spaceAfter=12)
    section_style = ParagraphStyle('SecTitle', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=13, leading=16, textColor=colors.HexColor("#1f77b4"), spaceBefore=12, spaceAfter=8)
    cxl_title_style = ParagraphStyle('CxlTitle', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=13, leading=16, textColor=colors.HexColor("#d62728"), spaceBefore=15, spaceAfter=8)
    
    table_header_style = ParagraphStyle('TH', fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.white, alignment=1)
    table_cell_style = ParagraphStyle('TC', fontName='Helvetica', fontSize=8, leading=10, textColor=colors.HexColor("#222222"), alignment=1)
    table_cell_bold = ParagraphStyle('TCB', fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.HexColor("#111111"), alignment=1)

    elements = []

    # Title Banner Blocks
    elements.append(Paragraph("Consolidated Performance & Operational Sales Matrix", title_style))
    elements.append(Paragraph(f"Active Portfolios Assessment | Executive Workbook Summary", subtitle_style))
    elements.append(Spacer(1, 5))

    # --- SECTION 1: ACCENTUATED OVERALL PERFORMANCE HIGHLIGHTS ---
    elements.append(Paragraph("1. Executive Summary & Core Financial Positions", section_style))
    
    perf_summary_data = [
        [Paragraph("<b>Financial KPI Metric</b>", table_header_style), Paragraph("<b>Consolidated Value Valuation</b>", table_header_style)],
        [Paragraph("TOTAL ULTIMATE NET COMBINED SALES", table_cell_bold), Paragraph(f"<b>{format_currency(grand_revenue)}</b>", table_cell_bold)],
        [Paragraph("TOTAL OVERALL VALUE DEDUCTIONS", table_cell_style), Paragraph(format_currency(total_deducted), table_cell_style)],
        [Paragraph("ACTIVE OPERATIONAL PORTFOLIOS", table_cell_style), Paragraph(f"{len(active_branches)} Branches Registered", table_cell_style)]
    ]
    perf_table = Table(perf_summary_data, colWidths=[3.75 * inch, 3.75 * inch])
    perf_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1f77b4")),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#d3d3d3")),
        ('BACKGROUND', (0,1), (-1,1), colors.HexColor("#eef4f8")),  # Highlight grand net row
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('TOPPADDING', (0,0), (-1,-1), 6),
    ]))
    elements.append(perf_table)
    elements.append(Spacer(1, 15))

    # --- SECTION 2: AGENCY AND PRODUCT MIX MATRIX WITH ACCENTUATED GRAND TOTALS ---
    elements.append(Paragraph("2. Agency and Product Mix Matrix", section_style))
    
    global_pivot = total_df.pivot_table(index="Agency", columns="Product_Type", values="NETMAINPRODUCT", aggfunc="sum", fill_value=0)
    for p_col in PRODUCT_ORDER:
        if p_col not in global_pivot.columns: global_pivot[p_col] = 0.0
    global_pivot = global_pivot[PRODUCT_ORDER].reindex(AGENCY_ORDER).fillna(0)

    headers = ["Agency", "FSP", "Pedestal", "Niche", "Others", "Total Net Sales", "Corp Net %"]
    table_data = [[Paragraph(h, table_header_style) for h in headers]]
    
    # Track metrics for structural ultimate column summing
    col_totals = {"FSP": 0.0, "Pedestal": 0.0, "Niche": 0.0, "Others": 0.0}
    
    for agency in AGENCY_ORDER:
        fsp = global_pivot.loc[agency, "FSP"]
        ped = global_pivot.loc[agency, "Pedestal"]
        nic = global_pivot.loc[agency, "Niche"]
        oth = global_pivot.loc[agency, "Others"]
        
        col_totals["FSP"] += fsp
        col_totals["Pedestal"] += ped
        col_totals["Niche"] += nic
        col_totals["Others"] += oth
        
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

    # ACCENTUATED ULTIMATE GRAND TOTAL ROW
    grand_row = [
        Paragraph("<b>GRAND TOTAL</b>", table_cell_bold),
        Paragraph(f"<b>{format_currency(col_totals['FSP'])}</b>", table_cell_bold),
        Paragraph(f"<b>{format_currency(col_totals['Pedestal'])}</b>", table_cell_bold),
        Paragraph(f"<b>{format_currency(col_totals['Niche'])}</b>", table_cell_bold),
        Paragraph(f"<b>{format_currency(col_totals['Others'])}</b>", table_cell_bold),
        Paragraph(f"<b>{format_currency(grand_revenue)}</b>", table_cell_bold),
        Paragraph("<b>100.00%</b>", table_cell_bold)
    ]
    table_data.append(grand_row)

    col_widths = [1.1*inch, 1.05*inch, 1.05*inch, 1.05*inch, 1.05*inch, 1.2*inch, 1.0*inch]
    global_table = Table(table_data, colWidths=col_widths, repeatRows=1)
    global_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1f77b4")),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#d3d3d3")),
        ('ROWBACKGROUNDS', (0,1), (-1,-2), [colors.white, colors.HexColor("#f9f9f9")]),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#eef4f8")), # Visual accentuation highlight formatting
        ('LINEABOVE', (0,-1), (-1,-1), 1.5, colors.HexColor("#1f77b4")),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('TOPPADDING', (0,0), (-1,-1), 5),
    ]))
    elements.append(global_table)
    elements.append(Spacer(1, 15))

    # --- SECTION 3: DETAILED BRANCH BREAKDOWNS ---
    elements.append(Paragraph("3. Operational Branch Breakdowns", section_style))
    for b_name, b_df in active_branches.items():
        b_total = b_df["NETMAINPRODUCT"].sum()
        
        # Enforce rule: No FSP tracking structures for LST and TLT
        current_branch_products = [p for p in PRODUCT_ORDER if not (b_name in ["LST", "TLT"] and p == "FSP")]
        
        b_pivot = b_df.pivot_table(index="Agency", columns="Product_Type", values="NETMAINPRODUCT", aggfunc="sum", fill_value=0)
        for col in current_branch_products:
            if col not in b_pivot.columns: b_pivot[col] = 0.0
        b_pivot = b_pivot[current_branch_products].reindex(AGENCY_ORDER).fillna(0)
        
        branch_elements = []
        branch_elements.append(Paragraph(f"Branch: {b_name} (Net Portfolio Revenue: {format_currency(b_total)})", ParagraphStyle('BName', fontName='Helvetica-Bold', fontSize=10, spaceBefore=6, spaceAfter=4, textColor=colors.HexColor("#333333"))))
        
        b_headers = ["Agency"] + current_branch_products + ["Branch Vol %", "Global Vol %"]
        b_table_data = [[Paragraph(bh, table_header_style) for bh in b_headers]]
        
        b_col_totals = {p: 0.0 for p in current_branch_products}
        for agency in AGENCY_ORDER:
            row_tot = 0.0
            row = [Paragraph(agency, table_cell_bold)]
            
            for prod in current_branch_products:
                val = b_pivot.loc[agency, prod]
                row_tot += val
                b_col_totals[prod] += val
                row.append(Paragraph(format_currency(val), table_cell_style))
                
            br_vol = (row_tot / max(b_total, 1)) * 100
            gl_vol = (row_tot / max(grand_revenue, 1)) * 100
            
            row.append(Paragraph(format_pct(br_vol), table_cell_style))
            row.append(Paragraph(format_pct(gl_vol), table_cell_style))
            b_table_data.append(row)
            
        # Accentuated Branch Total Row
        b_grand_row = [Paragraph("<b>TOTAL</b>", table_cell_bold)]
        for prod in current_branch_products:
            b_grand_row.append(Paragraph(f"<b>{format_currency(b_col_totals[prod])}</b>", table_cell_bold))
        b_grand_row.append(Paragraph(f"<b>{format_pct((b_total/max(b_total,1))*100)}</b>", table_cell_bold))
        b_grand_row.append(Paragraph(f"<b>{format_pct((b_total/max(grand_revenue,1))*100)}</b>", table_cell_bold))
        b_table_data.append(b_grand_row)
            
        if b_name in ["LST", "TLT"]:
            b_col_widths = [1.1*inch, 1.3*inch, 1.3*inch, 1.3*inch, 1.25*inch, 1.25*inch]
        else:
            b_col_widths = [1.1*inch, 1.05*inch, 1.05*inch, 1.05*inch, 1.05*inch, 1.2*inch, 1.0*inch]
            
        branch_table = Table(b_table_data, colWidths=b_col_widths, repeatRows=1)
        branch_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#2ca02c")),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#e2e2e2")),
            ('ROWBACKGROUNDS', (0,1), (-1,-2), [colors.white, colors.HexColor("#fcfcfc")]),
            ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#ebf7eb")),
            ('LINEABOVE', (0,-1), (-1,-1), 1.2, colors.HexColor("#2ca02c")),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('TOPPADDING', (0,0), (-1,-1), 4),
        ]))
        branch_elements.append(branch_table)
        
        # --- SUB-SECTION: MONTHLY TREND WITHOUT PRODUCT MIX ---
        branch_elements.append(Spacer(1, 4))
        b_month_df = b_df.groupby("Month")["NETMAINPRODUCT"].sum().reset_index()
        b_month_df = sort_by_month(b_month_df, "Month")
        
        bm_headers = ["Month Cycle Period", "Net Branch Revenue", "Branch Allocation %"]
        bm_table_data = [[Paragraph(bmh, table_header_style) for bmh in bm_headers]]
        
        for _, m_row in b_month_df.iterrows():
            m_val = m_row["NETMAINPRODUCT"]
            m_pct = (m_val / max(b_total, 1)) * 100
            bm_table_data.append([
                Paragraph(str(m_row["Month"]), table_cell_style),
                Paragraph(format_currency(m_val), table_cell_style),
                Paragraph(format_pct(m_pct), table_cell_style)
            ])
            
        bm_table = Table(bm_table_data, colWidths=[2.5 * inch, 2.5 * inch, 2.5 * inch])
        bm_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#7f7f7f")),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#e2e2e2")),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
            ('TOPPADDING', (0,0), (-1,-1), 3),
        ]))
        branch_elements.append(bm_table)
        branch_elements.append(Spacer(1, 12))
        
        elements.append(KeepTogether(branch_elements))

    # --- SECTION 4: SALES CANCELLATION MATRIX ---
    cxl_block = []
    cxl_block.append(Paragraph(f"4. Sales Cancellation Breakdown (Total Deductions: {format_currency(total_deducted)})", cxl_title_style))
    
    if valid_cxl_df.empty:
        cxl_block.append(Paragraph("<i>No cancellation rows registered inside this fiscal cycle workbook.</i>", table_cell_style))
    else:
        cxl_pivot = valid_cxl_df.pivot_table(index="Agency", columns="Product_Type", values="Sales_Amount", aggfunc="sum", fill_value=0)
        for p_col in PRODUCT_ORDER:
            if p_col not in cxl_pivot.columns: cxl_pivot[p_col] = 0.0
        cxl_pivot = cxl_pivot[PRODUCT_ORDER].reindex(AGENCY_ORDER).fillna(0)
        
        cxl_headers = ["Agency Matrix", "FSP", "Pedestal", "Niche", "Others", "Total Rev Voided", "Contribution %"]
        cxl_table_data = [[Paragraph(ch, table_header_style) for ch in cxl_headers]]
        
        c_col_totals = {"FSP": 0.0, "Pedestal": 0.0, "Niche": 0.0, "Others": 0.0}
        for agency in AGENCY_ORDER:
            fsp_c = cxl_pivot.loc[agency, "FSP"]
            ped_c = cxl_pivot.loc[agency, "Pedestal"]
            nic_c = cxl_pivot.loc[agency, "Niche"]
            oth_c = cxl_pivot.loc[agency, "Others"]
            
            c_col_totals["FSP"] += fsp_c
            c_col_totals["Pedestal"] += ped_c
            c_col_totals["Niche"] += nic_c
            c_col_totals["Others"] += oth_c
            
            row_cxl_tot = fsp_c + ped_c + nic_c + oth_c
            risk_pct = (row_cxl_tot / max(total_deducted, 1)) * 100
            
            row_c = [
                Paragraph(agency, table_cell_bold),
                Paragraph(format_currency(fsp_c), table_cell_style),
                Paragraph(format_currency(ped_c), table_cell_style),
                Paragraph(format_currency(nic_c), table_cell_style),
                Paragraph(format_currency(oth_c), table_cell_style),
                Paragraph(format_currency(row_cxl_tot), table_cell_bold),
                Paragraph(format_pct(risk_pct), table_cell_style)
            ]
            cxl_table_data.append(row_c)
            
        # Accentuated Cancellation Grand Row
        c_grand_row = [
            Paragraph("<b>TOTAL CANCELLATIONS</b>", table_cell_bold),
            Paragraph(f"<b>{format_currency(c_col_totals['FSP'])}</b>", table_cell_bold),
            Paragraph(f"<b>{format_currency(c_col_totals['Pedestal'])}</b>", table_cell_bold),
            Paragraph(f"<b>{format_currency(c_col_totals['Niche'])}</b>", table_cell_bold),
            Paragraph(f"<b>{format_currency(c_col_totals['Others'])}</b>", table_cell_bold),
            Paragraph(f"<b>{format_currency(total_deducted)}</b>", table_cell_bold),
            Paragraph("<b>100.00%</b>", table_cell_bold)
        ]
        cxl_table_data.append(c_grand_row)
            
        cxl_table = Table(cxl_table_data, colWidths=col_widths, repeatRows=1)
        cxl_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#d62728")),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#f4c2c2")),
            ('ROWBACKGROUNDS', (0,1), (-1,-2), [colors.white, colors.HexColor("#fff5f5")]),
            ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#fdebe8")),
            ('LINEABOVE', (0,-1), (-1,-1), 1.5, colors.HexColor("#d62728")),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
            ('TOPPADDING', (0,0), (-1,-1), 5),
        ]))
        cxl_block.append(cxl_table)
        
    elements.append(KeepTogether(cxl_block))

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

        # Dynamically map the 6th Excel column (Column F -> index 5)
        df['Month'] = pd.to_datetime(df.iloc[:, 5], errors='coerce').dt.strftime('%B')
        df['Month'] = df['Month'].fillna('Unknown')

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
total_deducted_amount = 0.0

if edited_cxl_df is not None and not edited_cxl_df.empty:
    valid_cxl_df = edited_cxl_df.dropna(subset=["Branch", "Agency", "Product_Type", "Sales_Amount", "Qty"]).copy()
    valid_cxl_df["Total_Deduction"] = valid_cxl_df["Sales_Amount"] * valid_cxl_df["Qty"]
    total_deducted_amount = valid_cxl_df["Total_Deduction"].sum()
    
    for _, item in valid_cxl_df.iterrows():
        row_deduction = item["Total_Deduction"]
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
    pdf_data = generate_pdf_report(grand_corporate_revenue, total_df, active_branches, valid_cxl_df, total_deducted_amount)
    st.sidebar.download_button(
        label="🔴 Download Executive PDF Report",
        data=pdf_data,
        file_name=f"Executive_Sales_Performance_Report.pdf",
        mime="application/pdf",
        use_container_width=True
    )
except Exception as err:
    st.sidebar.error(f"PDF Compiler Standby: {err}")

# -------------------- MAIN DASHBOARD VISUALIZATIONS (STREAMLIT SCREEN) --------------------
tab_titles = ["Consolidated Overview"] + [f"{name} Branch" for name in active_branches.keys()]
tabs = st.tabs(tab_titles)

with tabs[0]:
    st.header("🏢 Unified Corporate Overview")
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Net Combined Sales", format_currency(grand_corporate_revenue))
    c2.metric("Active Operating Branches", len(active_branches))
    c3.metric("Total Value Deducted", format_currency(total_deducted_amount))
    
    st.write("---")
    st.subheader("Overall Net Performance by Agency Hierarchy (Visual Chart)")
    if not total_df.empty:
        global_agency_prod = total_df.groupby(["Agency", "Product_Type"])["NETMAINPRODUCT"].sum().reset_index()
        global_agency_prod = sort_by_corporate_hierarchy(global_agency_prod, "Agency")
        fig_gap = px.bar(global_agency_prod, x="Agency", y="NETMAINPRODUCT", color="Product_Type", color_discrete_map=product_colours, barmode="stack", text=global_agency_prod["NETMAINPRODUCT"].apply(format_chart_label))
        st.plotly_chart(fig_gap, use_container_width=True)
        
        st.markdown("**Agency and Product Mix Matrix Summary Table**")
        ui_global_pivot = total_df.pivot_table(index="Agency", columns="Product_Type", values="NETMAINPRODUCT", aggfunc="sum", fill_value=0)
        ui_global_pivot = ui_global_pivot.reindex(index=AGENCY_ORDER, columns=PRODUCT_ORDER, fill_value=0)
        ui_global_pivot["Total Revenue"] = ui_global_pivot.sum(axis=1)
        st.dataframe(format_currency_df(ui_global_pivot.reset_index(), PRODUCT_ORDER + ["Total Revenue"]), use_container_width=True, hide_index=True)
        
        st.write("---")
        st.subheader("Overall Portfolio Share Mix (Visual Chart)")
        prod_sum = total_df.groupby("Product_Type")["NETMAINPRODUCT"].sum().reset_index()
        fig_p = px.pie(prod_sum, names="Product_Type", values="NETMAINPRODUCT", hole=0.3, color="Product_Type", color_discrete_map=product_colours)
        st.plotly_chart(fig_p, use_container_width=True)
        
        st.markdown("**Data Table: Portfolio Share Breakdown**")
        prod_sum_display = prod_sum.copy()
        prod_sum_display["Contribution %"] = (prod_sum_display["NETMAINPRODUCT"] / max(grand_corporate_revenue, 1)) * 100
        prod_sum_display["Contribution %"] = prod_sum_display["Contribution %"].apply(format_pct)
        st.dataframe(format_currency_df(prod_sum_display, ["NETMAINPRODUCT"]), use_container_width=True, hide_index=True)
    else:
        st.warning("No operating margins left to display.")

for idx, (b_name, b_df) in enumerate(active_branches.items(), start=1):
    with tabs[idx]:
        st.header(f"📍 Operational Analysis: {b_name} Branch")
        if not b_df.empty:
            # Layout splitting: Product Mix on left, New Monthly Metric Breakdown on right
            left_col, right_col = st.columns([3, 2])
            
            with left_col:
                st.subheader("Branch Performance Breakdown (Visual Chart)")
                branch_agency_prod = b_df.groupby(["Agency", "Product_Type"])["NETMAINPRODUCT"].sum().reset_index()
                if b_name in ["LST", "TLT"]:
                    branch_agency_prod = branch_agency_prod[branch_agency_prod["Product_Type"] != "FSP"]
                    current_branch_products = [p for p in PRODUCT_ORDER if p != "FSP"]
                else:
                    current_branch_products = PRODUCT_ORDER

                branch_agency_prod = sort_by_corporate_hierarchy(branch_agency_prod, "Agency")
                fig_br_ap = px.bar(branch_agency_prod, x="Agency", y="NETMAINPRODUCT", color="Product_Type", color_discrete_map=product_colours, barmode="stack", text=branch_agency_prod["NETMAINPRODUCT"].apply(format_chart_label))
                st.plotly_chart(fig_br_ap, use_container_width=True)
                
                st.markdown(f"**Data Table: {b_name} Product Matrix**")
                ui_branch_pivot = b_df.pivot_table(index="Agency", columns="Product_Type", values="NETMAINPRODUCT", aggfunc="sum", fill_value=0)
                ui_branch_pivot = ui_branch_pivot.reindex(index=AGENCY_ORDER, columns=current_branch_products, fill_value=0)
                ui_branch_pivot["Branch Total"] = ui_branch_pivot.sum(axis=1)
                st.dataframe(format_currency_df(ui_branch_pivot.reset_index(), current_branch_products + ["Branch Total"]), use_container_width=True, hide_index=True)

            with right_col:
                st.subheader("Monthly Revenue Net Run-Rate")
                branch_monthly = b_df.groupby("Month")["NETMAINPRODUCT"].sum().reset_index()
                branch_monthly = sort_by_month(branch_monthly, "Month")
                
                # Visual continuous line chart representing the monthly net run-rate trend
                fig_month = px.line(branch_monthly, x="Month", y="NETMAINPRODUCT", markers=True, text=branch_monthly["NETMAINPRODUCT"].apply(format_chart_label), labels={"NETMAINPRODUCT": "Net Sales ($)"})
                fig_month.update_traces(textposition="top center", line_color="#1f77b4")
                st.plotly_chart(fig_month, use_container_width=True)
                
                st.markdown(f"**Data Table: Chronological Cycles Overview**")
                st.dataframe(format_currency_df(branch_monthly, ["NETMAINPRODUCT"]), use_container_width=True, hide_index=True)
