import streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
import datetime

# -------------------- PAGE CONFIG --------------------
st.set_page_config(page_title="Multi-Branch Sales Dashboard", layout="wide")

st.title("📊 Sales Dashboard - Multi-Branch Analysis")
st.markdown("Upload Excel files for **CCK**, **LST**, and **TLT** branches, log cancellations, and export structural PDF audit reports.")

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

# -------------------- PDF EXPORT FUNCTION --------------------
class CorporateReportPDF(FPDF):
    def header(self):
        self.set_font("Helvetica", "B", 12)
        self.set_text_color(100, 100, 100)
        self.cell(0, 10, "CONSOLIDATED CORPORATE PERFORMANCE REPORT", ln=True, align="L")
        self.line(10, 18, 200, 18)
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f"Generated on {datetime.date.today().strftime('%d %b %Y')} | Page {self.page_no()}", align="C")

def generate_pdf_report(grand_revenue, raw_df, valid_cxl_df, active_branches):
    pdf = CorporateReportPDF()
    pdf.add_page()
    
    # Title Block
    pdf.set_font("Helvetica", "B", 22)
    pdf.set_text_color(31, 119, 180)
    pdf.cell(0, 15, "Unified Executive Sales Summary", ln=True)
    pdf.ln(2)
    
    # Financial Overview Cards (Table-Style Layout)
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_fill_color(240, 242, 246)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(60, 8, "Total Net Combined Sales", border=1, fill=True)
    pdf.cell(65, 8, "Confirmed Transactions Count", border=1, fill=True)
    pdf.cell(65, 8, "Active Operating Branches", border=1, fill=True, ln=True)
    
    pdf.set_font("Helvetica", "", 12)
    pdf.cell(60, 10, format_currency(grand_revenue), border=1)
    pdf.cell(65, 10, f"{len(raw_df):,} orders", border=1)
    pdf.cell(65, 10, f"{len(active_branches)} branches", border=1, ln=True)
    pdf.ln(8)
    
    # Cancellations Audit Log Section
    pdf.set_font("Helvetica", "B", 14)
    pdf.set_text_color(214, 39, 40)
    pdf.cell(0, 10, "Applied Deductions & Cancellations Registry", ln=True)
    
    if valid_cxl_df is not None and not valid_cxl_df.empty:
        pdf.set_font("Helvetica", "B", 10)
        pdf.set_fill_color(230, 230, 230)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(30, 7, "Branch", border=1, fill=True)
        pdf.cell(35, 7, "Agency", border=1, fill=True)
        pdf.cell(45, 7, "Product Type", border=1, fill=True)
        pdf.cell(20, 7, "Qty", border=1, fill=True)
        pdf.cell(30, 7, "Unit Amt", border=1, fill=True)
        pdf.cell(30, 7, "Total Drop", border=1, fill=True, ln=True)
        
        pdf.set_font("Helvetica", "", 10)
        for _, row in valid_cxl_df.iterrows():
            total_drop = row['Sales_Amount'] * row['Qty']
            pdf.cell(30, 6, str(row['Branch']), border=1)
            pdf.cell(35, 6, str(row['Agency']), border=1)
            pdf.cell(45, 6, str(row['Product_Type']), border=1)
            pdf.cell(20, 6, str(row['Qty']), border=1, align="C")
            pdf.cell(30, 6, format_currency(row['Sales_Amount']), border=1, align="R")
            pdf.cell(30, 6, format_currency(total_drop), border=1, align="R", ln=True)
    else:
        pdf.set_font("Helvetica", "I", 11)
        pdf.set_text_color(100, 100, 100)
        pdf.cell(0, 8, "No corporate cancellations logged for this specific matrix sequence run.", ln=True)
    
    pdf.ln(10)
    
    # Branch Breakdown Sections
    pdf.set_font("Helvetica", "B", 14)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 10, "Branch-Specific Net Breakdown Summary", ln=True)
    pdf.ln(2)
    
    for b_name, b_df in active_branches.items():
        pdf.set_font("Helvetica", "B", 12)
        pdf.set_text_color(44, 160, 44)
        pdf.cell(0, 8, f"Branch Location: {b_name} (Net Revenue: {format_currency(b_df['NETMAINPRODUCT'].sum())})", ln=True)
        
        # Build quick text pivot mapping matrix metrics inside PDF
        pivot = b_df.pivot_table(index="Agency", columns="Product_Type", values="NETMAINPRODUCT", aggfunc="sum", fill_value=0)
        for col in PRODUCT_ORDER:
            if col not in pivot.columns: pivot[col] = 0.0
            
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_fill_color(245, 245, 245)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(35, 6, "Agency", border=1, fill=True)
        pdf.cell(30, 6, "FSP", border=1, fill=True, align="R")
        pdf.cell(30, 6, "Pedestal", border=1, fill=True, align="R")
        pdf.cell(30, 6, "Niche", border=1, fill=True, align="R")
        pdf.cell(30, 6, "Others", border=1, fill=True, align="R")
        pdf.cell(35, 6, "Total Revenue", border=1, fill=True, align="R", ln=True)
        
        pdf.set_font("Helvetica", "", 9)
        for agency in AGENCY_ORDER:
            if agency in pivot.index:
                fsp = pivot.loc[agency, "FSP"]
                ped = pivot.loc[agency, "Pedestal"]
                nic = pivot.loc[agency, "Niche"]
                oth = pivot.loc[agency, "Others"]
                row_tot = fsp + ped + nic + oth
                
                pdf.cell(35, 6, agency, border=1)
                pdf.cell(30, 6, format_currency(fsp), border=1, align="R")
                pdf.cell(30, 6, format_currency(ped), border=1, align="R")
                pdf.cell(30, 6, format_currency(nic), border=1, align="R")
                pdf.cell(30, 6, format_currency(oth), border=1, align="R")
                pdf.cell(35, 6, format_currency(row_tot), border=1, align="R", ln=True)
        pdf.ln(6)
        
    return pdf.output()

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
total_deducted_amount = 0.0
valid_cxl_df = pd.DataFrame()

if edited_cxl_df is not None and not edited_cxl_df.empty:
    valid_cxl_df = edited_cxl_df.dropna(subset=["Branch", "Agency", "Product_Type", "Sales_Amount", "Qty"])
    for _, item in valid_cxl_df.iterrows():
        row_deduction = item["Sales_Amount"] * item["Qty"]
        if row_deduction <= 0: continue
        total_deducted_amount += row_deduction
        
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

# -------------------- PDF EXPORT BUTTON INTEGRATION --------------------
st.sidebar.write("---")
st.sidebar.header("🖨️ Document Exports")
try:
    generated_pdf_bytes = generate_pdf_report(grand_corporate_revenue, total_df, valid_cxl_df, active_branches)
    st.sidebar.download_button(
        label="📥 Download Consolidated PDF Report",
        data=generated_pdf_bytes,
        file_name=f"Consolidated_Performance_Report_{datetime.date.today().strftime('%Y%m%d')}.pdf",
        mime="application/pdf",
        use_container_width=True
    )
except Exception as pdf_err:
    st.sidebar.error(f"PDF Compiler Standby: {pdf_err}")

# -------------------- MAIN DASHBOARD VISUALIZATIONS --------------------
tab_titles = ["Consolidated Overview"] + [f"{name} Branch" for name in active_branches.keys()]
tabs = st.tabs(tab_titles)

with tabs[0]:
    st.header("🏢 Unified Corporate Overview")
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Net Combined Sales", format_currency(grand_corporate_revenue))
    c2.metric("Active Operating Branches", len(active_branches))
    c3.metric("Net Confirmed Transactions", f"{len(total_df):,}", delta=f"-{format_currency(total_deducted_amount)} Capped" if total_deducted_amount > 0 else None, delta_color="inverse")

    st.write("---")
    st.subheader("Overall Performance by Agency Hierarchy")
    if not total_df.empty:
        global_agency_prod = total_df.groupby(["Agency", "Product_Type"])["NETMAINPRODUCT"].sum().reset_index()
        global_agency_prod = sort_by_corporate_hierarchy(global_agency_prod, "Agency")
        fig_gap = px.bar(global_agency_prod, x="Agency", y="NETMAINPRODUCT", color="Product_Type", color_discrete_map=product_colours, barmode="stack", text=global_agency_prod["NETMAINPRODUCT"].apply(format_chart_label))
        st.plotly_chart(fig_gap, use_container_width=True)

        global_pivot = total_df.pivot_table(index="Agency", columns="Product_Type", values="NETMAINPRODUCT", aggfunc="sum", fill_value=0)
        for p_col in PRODUCT_ORDER:
            if p_col not in global_pivot.columns: global_pivot[p_col] = 0.0
        global_pivot = global_pivot[PRODUCT_ORDER]
        global_pivot["Total Contribution ($)"] = global_pivot.sum(axis=1)
        global_pivot["Overall Corporate Contribution %"] = (global_pivot["Total Contribution ($)"] / max(grand_corporate_revenue, 1)) * 100
        global_pivot = global_pivot.reindex(AGENCY_ORDER).fillna(0)

        st.markdown("**Data Matrix: Overall Agency Product Mix**")
        display_global_pivot = format_currency_df(global_pivot, ["FSP", "Pedestal", "Niche", "Others", "Total Contribution ($)"])
        display_global_pivot["Overall Corporate Contribution %"] = display_global_pivot["Overall Corporate Contribution %"].apply(format_pct)
        st.dataframe(display_global_pivot, use_container_width=True)
    else:
        st.warning("No operating margins left to display.")

for idx, (b_name, b_df) in enumerate(active_branches.items(), start=1):
    with tabs[idx]:
        st.header(f"📍 Operational Analysis: {b_name} Branch")
        b_total = b_df["NETMAINPRODUCT"].sum()
        mc1, mc2, mc3 = st.columns(3)
        mc1.metric(f"{b_name} Net Revenue", format_currency(b_total))
        mc2.metric("Active Local Agencies", b_df["Agency"].nunique())
        mc3.metric("Volume of Line Orders", f"{len(b_df):,}")

        if not b_df.empty:
            branch_agency_prod = b_df.groupby(["Agency", "Product_Type"])["NETMAINPRODUCT"].sum().reset_index()
            branch_agency_prod = sort_by_corporate_hierarchy(branch_agency_prod, "Agency")
            fig_br_ap = px.bar(branch_agency_prod, x="Agency", y="NETMAINPRODUCT", color="Product_Type", color_discrete_map=product_colours, barmode="stack", text=branch_agency_prod["NETMAINPRODUCT"].apply(format_chart_label))
            st.plotly_chart(fig_br_ap, use_container_width=True)
