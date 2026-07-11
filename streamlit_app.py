import streamlit as st
import pandas as pd
import plotly.express as px
import traceback

# -------------------- PAGE CONFIG --------------------
st.set_page_config(page_title="Multi-Branch Sales Dashboard", layout="wide")

st.title("📊 Sales Dashboard - Multi-Branch Analysis")
st.markdown("Upload Excel files for **CCK**, **LST**, and **TLT** branches, and log manual cancellations to view updated insights.")

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

# -------------------- SESSION STATE INITIALISATION --------------------
# This keeps track of manual cancellations across app re-runs
if "cancellation_registry" not in st.session_state:
    st.session_state.cancellation_registry = []

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

# -------------------- DATA PROCESSING ENGINES --------------------
def process_branch_file(uploaded_file, branch_name):
    """Extract confirmed sales data from a branch Excel file."""
    try:
        df_raw = pd.read_excel(uploaded_file, header=None)
        header_row_idx = None
        for idx, row in df_raw.iterrows():
            if row.astype(str).str.contains("FILE_NO").any():
                header_row_idx = idx
                break
        if header_row_idx is None:
            return None

        df = pd.read_excel(uploaded_file, header=header_row_idx)
        df.columns = df.columns.str.strip()

        required_cols = ["STATUS", "NETMAINPRODUCT", "CBDD_NAME", "BDD_NAME", "PRODUCT_CODE"]
        if any(col not in df.columns for col in required_cols):
            return None

        df["NETMAINPRODUCT"] = pd.to_numeric(df["NETMAINPRODUCT"], errors="coerce")
        df_confirmed = df[df["STATUS"].str.upper() == "CONFIRM"].copy()
        df_confirmed = df_confirmed[df_confirmed["NETMAINPRODUCT"].notna() & (df_confirmed["NETMAINPRODUCT"] != 0)]

        if df_confirmed.empty:
            return pd.DataFrame()

        # Agency mapping
        agency_rename = {
            "FU GUI SERVICES": "FGY",
            "ZENBOX PTE LTD": "ZB",
            "APG ADVISORY PTE. LTD.": "APG",
            "SINGAPORE LIFESTYLE ASSOCIATES PTE LTD.": "SLA",
            "JF LIFE CONSULTANT PTE LTD": "JFL",
        }

        def get_agency(row):
            cbd = row.get("CBDD_NAME", "")
            if pd.notna(cbd) and str(cbd).strip() != "":
                raw = str(cbd).strip().upper()
            else:
                bdd = row.get("BDD_NAME", "")
                raw = str(bdd).strip().upper() if pd.notna(bdd) and str(bdd).strip() != "" else "Others"
            for key, value in agency_rename.items():
                if key in raw:
                    return value
            return "Others"

        df_confirmed["Agency"] = df_confirmed.apply(get_agency, axis=1)

        def get_product_type(code):
            if pd.isna(code):
                return "Others"
            code_str = str(code).strip().upper()
            if code_str == "P":
                return "FSP"
            elif code_str == "TABLET":
                return "Pedestal"
            elif code_str == "URN":
                return "Niche"
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

# -------------------- SIDEBAR FILE UPLOADERS --------------------
st.sidebar.header("📁 Branch File Uploaders")
cck_file = st.sidebar.file_uploader("Upload CCK Branch Excel", type=["xlsx", "xls"])
lst_file = st.sidebar.file_uploader("Upload LST Branch Excel", type=["xlsx", "xls"])
tlt_file = st.sidebar.file_uploader("Upload TLT Branch Excel", type=["xlsx", "xls"])

# -------------------- SIDEBAR CANCELLATION INTERFACE --------------------
st.sidebar.write("---")
st.sidebar.header("🛑 Manual Cancellation Registry")

with st.sidebar.form(key="cancellation_form", clear_on_submit=True):
    cxl_branch = st.selectbox("Target Branch", BRANCH_OPTIONS)
    cxl_agency = st.selectbox("Target Agency", AGENCY_ORDER)
    cxl_product = st.selectbox("Target Product", PRODUCT_ORDER)
    cxl_amount = st.number_input("Sales Amount per Item ($)", min_value=0.0, step=500.0, format="%.2f")
    cxl_qty = st.number_input("Cancellation QTY", min_value=1, step=1, value=1)
    
    submit_cxl = st.form_submit_button("Log & Submit Cancellation")
    
    if submit_cxl:
        total_deduction = cxl_amount * cxl_qty
        if total_deduction > 0:
            st.session_state.cancellation_registry.append({
                "Branch": cxl_branch,
                "Agency": cxl_agency,
                "Product_Type": cxl_product,
                "Qty": cxl_qty,
                "Amount": cxl_amount,
                "Total_Deduction": total_deduction
            })
            st.success(f"Log Successful: Deducting {format_currency(total_deduction)} from {cxl_branch} → {cxl_agency} → {cxl_product}!")
        else:
            st.warning("Please enter a valid Sales Amount higher than $0.")

# Clear Button to wipe logged records if needed
if st.session_state.cancellation_registry:
    if st.sidebar.button("Wipe & Reset All Logged Cancellations"):
        st.session_state.cancellation_registry = []
        st.sidebar.success("Cancellation history cleared!")
        st.rerun()

# -------------------- PROCESS BRANCH FILES --------------------
branch_dfs = {}
if cck_file:
    branch_dfs["CCK"] = process_branch_file(cck_file, "CCK")
if lst_file:
    branch_dfs["LST"] = process_branch_file(lst_file, "LST")
if tlt_file:
    branch_dfs["TLT"] = process_branch_file(tlt_file, "TLT")

active_branches = {k: v for k, v in branch_dfs.items() if v is not None and not v.empty}

if not active_branches:
    st.info("👋 Welcome! Please upload at least one valid Branch Excel data sheet via the sidebar to initialize the dashboard panels.")
    st.stop()

# Consolidate raw uploaded data
total_df = pd.concat(active_branches.values(), ignore_index=True)

# -------------------- APPLY ATTRIBUTE-BASED DEDUCTIONS --------------------
total_deducted_amount = 0.0

if st.session_state.cancellation_registry:
    for item in st.session_state.cancellation_registry:
        total_deducted_amount += item["Total_Deduction"]
        
        # Locate the structural intersections in our primary dataframe pool
        mask = (
            (total_df["Branch"] == item["Branch"]) & 
            (total_df["Agency"] == item["Agency"]) & 
            (total_df["Product_Type"] == item["Product_Type"])
        )
        
        matching_rows = total_df[mask]
        
        if not matching_rows.empty:
            # Sort matching rows lowest-value first to systematically chip away at total deduction target
            matching_indices = matching_rows.sort_values(by="NETMAINPRODUCT").index
            deduction_left = item["Total_Deduction"]
            
            for idx in matching_indices:
                current_val = total_df.at[idx, "NETMAINPRODUCT"]
                
                if current_val <= deduction_left:
                    # Wipe out/zero this row entirely since its value is less than or equal to what remains
                    deduction_left -= current_val
                    total_df.at[idx, "NETMAINPRODUCT"] = 0.0
                else:
                    # Subtract the remaining deduction from this row's ledger entry and terminate row loop
                    total_df.at[idx, "NETMAINPRODUCT"] = current_val - deduction_left
                    deduction_left = 0.0
                    break
    
    # Filter out records that were reduced down to zero to clean up downstream visualization matrix counts
    total_df = total_df[total_df["NETMAINPRODUCT"] != 0].copy()

    # Re-slice updated modifications back into individual active branch components
    for name in list(active_branches.keys()):
        active_branches[name] = total_df[total_df["Branch"] == name].copy()

grand_corporate_revenue = total_df['NETMAINPRODUCT'].sum()

# -------------------- SHOW ACTIVE MANUAL CANCELLATIONS TABLE --------------------
if st.session_state.cancellation_registry:
    with st.expander("📄 Active Cancellation Audit Registry Matrix", expanded=False):
        audit_df = pd.DataFrame(st.session_state.cancellation_registry)
        display_audit = format_currency_df(audit_df, ["Amount", "Total_Deduction"])
        st.dataframe(display_audit, use_container_width=True, hide_index=True)

# -------------------- BUILD TABS --------------------
tab_titles = ["Consolidated Overview"] + [f"{name} Branch" for name in active_branches.keys()]
tabs = st.tabs(tab_titles)

# ==================== TAB 0: CONSOLIDATED OVERVIEW ====================
with tabs[0]:
    st.header("🏢 Unified Corporate Overview")

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Net Combined Sales", format_currency(grand_corporate_revenue))
    c2.metric("Active Operating Branches", len(active_branches))
    
    if total_deducted_amount > 0:
        c3.metric("Net Confirmed Transactions", f"{len(total_df):,}", delta=f"-{format_currency(total_deducted_amount)} Cancelled", delta_color="inverse")
    else:
        c3.metric("Total Confirmed Transactions", f"{len(total_df):,}")

    st.write("---")

    st.subheader("Overall Performance by Agency Hierarchy")
    if not total_df.empty:
        global_agency_prod = total_df.groupby(["Agency", "Product_Type"])["NETMAINPRODUCT"].sum().reset_index()
        global_agency_prod = sort_by_corporate_hierarchy(global_agency_prod, "Agency")

        fig_gap = px.bar(
            global_agency_prod,
            x="Agency",
            y="NETMAINPRODUCT",
            color="Product_Type",
            color_discrete_map=product_colours,
            barmode="stack",
            text=global_agency_prod["NETMAINPRODUCT"].apply(format_chart_label)
        )
        fig_gap.update_layout(xaxis={'categoryorder': 'array', 'categoryarray': AGENCY_ORDER})
        st.plotly_chart(fig_gap, use_container_width=True)

        # Pivot table
        global_pivot = total_df.pivot_table(
            index="Agency",
            columns="Product_Type",
            values="NETMAINPRODUCT",
            aggfunc="sum",
            fill_value=0
        )
        for p_col in PRODUCT_ORDER:
            if p_col not in global_pivot.columns:
                global_pivot[p_col] = 0.0
        global_pivot = global_pivot[PRODUCT_ORDER]
        global_pivot["Total Contribution ($)"] = global_pivot.sum(axis=1)
        global_pivot["Overall Corporate Contribution %"] = (global_pivot["Total Contribution ($)"] / max(grand_corporate_revenue, 1)) * 100
        global_pivot = global_pivot.reindex(AGENCY_ORDER).fillna(0)

        st.markdown("**Data Matrix: Overall Agency Product Mix**")
        display_global_pivot = format_currency_df(
            global_pivot,
            ["FSP", "Pedestal", "Niche", "Others", "Total Contribution ($)"]
        )
        display_global_pivot["Overall Corporate Contribution %"] = display_global_pivot["Overall Corporate Contribution %"].apply(format_pct)
        st.dataframe(display_global_pivot, use_container_width=True)

        st.write("---")
        st.subheader("Overall Performance by Product Portfolio")
        prod_sum = total_df.groupby("Product_Type")["NETMAINPRODUCT"].sum().reset_index()
        fig_p = px.pie(
            prod_sum,
            names="Product_Type",
            values="NETMAINPRODUCT",
            hole=0.3,
            color="Product_Type",
            color_discrete_map=product_colours
        )
        st.plotly_chart(fig_p, use_container_width=True)
    else:
        st.warning("No sales data available. Logged deductions match or exceed complete branch turnover volume.")

# ==================== DYNAMIC BRANCH TABS ====================
for idx, (b_name, b_df) in enumerate(active_branches.items(), start=1):
    with tabs[idx]:
        st.header(f"📍 Operational Analysis: {b_name} Branch")
        b_total = b_df["NETMAINPRODUCT"].sum()

        mc1, mc2, mc3 = st.columns(3)
        mc1.metric(f"{b_name} Net Revenue", format_currency(b_total))
        mc2.metric("Active Local Agencies", b_df["Agency"].nunique())
        mc3.metric("Volume of Line Orders", f"{len(b_df):,}")

        st.subheader(f"Agency Performance & Product Mix in {b_name}")
        if not b_df.empty:
            branch_agency_prod = b_df.groupby(["Agency", "Product_Type"])["NETMAINPRODUCT"].sum().reset_index()
            branch_agency_prod = sort_by_corporate_hierarchy(branch_agency_prod, "Agency")

            fig_br_ap = px.bar(
                branch_agency_prod,
                x="Agency",
                y="NETMAINPRODUCT",
                color="Product_Type",
                color_discrete_map=product_colours,
                barmode="stack",
                text=branch_agency_prod["NETMAINPRODUCT"].apply(format_chart_label)
            )
            fig_br_ap.update_layout(xaxis={'categoryorder': 'array', 'categoryarray': AGENCY_ORDER})
            st.plotly_chart(fig_br_ap, use_container_width=True)

            branch_pivot = b_df.pivot_table(
                index="Agency",
                columns="Product_Type",
                values="NETMAINPRODUCT",
                aggfunc="sum",
                fill_value=0
            )
            for p_col in PRODUCT_ORDER:
                if p_col not in branch_pivot.columns:
                    branch_pivot[p_col] = 0.0
            branch_pivot = branch_pivot[PRODUCT_ORDER]
            branch_pivot["Total Branch Sales ($)"] = branch_pivot.sum(axis=1)
            branch_pivot["Agency Contribution in Branch %"] = (branch_pivot["Total Branch Sales ($)"] / max(b_total, 1)) * 100
            branch_pivot["Agency Branch Contribution in Overall %"] = (branch_pivot["Total Branch Sales ($)"] / max(grand_corporate_revenue, 1)) * 100
            branch_pivot = branch_pivot.reindex(AGENCY_ORDER).fillna(0)

            st.markdown(f"**Data Matrix: {b_name} Agency Performance Summary**")
            display_branch_pivot = format_currency_df(
                branch_pivot,
                ["FSP", "Pedestal", "Niche", "Others", "Total Branch Sales ($)"]
            )
            display_branch_pivot["Agency Contribution in Branch %"] = display_branch_pivot["Agency Contribution in Branch %"].apply(format_pct)
            display_branch_pivot["Agency Branch Contribution in Overall %"] = display_branch_pivot["Agency Branch Contribution in Overall %"].apply(format_pct)
            st.dataframe(display_branch_pivot, use_container_width=True)

            st.write("---")
            st.markdown("**Raw Net Branch Ledger Records**")
            display_raw = b_df[["FILE_NO", "Agency", "Product_Type", "NETMAINPRODUCT"]].copy()
            display_raw = sort_by_corporate_hierarchy(display_raw, "Agency")
            display_raw = format_currency_df(display_raw, ["NETMAINPRODUCT"])
            st.dataframe(display_raw, use_container_width=True, hide_index=True)
        else:
            st.info("All metrics for this specific branch were completely zeroed out by applied cancellations.")
