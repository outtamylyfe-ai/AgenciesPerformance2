import streamlit as st
import pandas as pd
import plotly.express as px
import traceback

# -------------------- PAGE CONFIG --------------------
st.set_page_config(page_title="Multi-Branch Sales Dashboard", layout="wide")

st.title("📊 Sales Dashboard - Multi-Branch Analysis")
st.markdown("Upload Excel files for **CCK**, **LST**, and **TLT** branches alongside **Cancellations** to view insights.")

# -------------------- CONSTANTS --------------------
AGENCY_ORDER = ["FGY", "ZB", "APG", "SLA", "JFL"]
PRODUCT_ORDER = ["FSP", "Pedestal", "Niche", "Others"]
product_colours = {"FSP": "#1f77b4", "Pedestal": "#ff7f0e", "Niche": "#2ca02c", "Others": "#d62728"}

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
            return raw

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
            df_confirmed["FILE_NO"] = df_confirmed["FILE_NO"].astype(str).str.strip()
        return df_confirmed

    except Exception as e:
        st.error(f"Error processing {branch_name} file: {e}")
        return None


def process_cancellation_file(uploaded_file):
    """Extract cancellation list matching solely on FILE_NO."""
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

        if "FILE_NO" not in df.columns:
            st.error("The Cancellation sheet must contain a 'FILE_NO' column.")
            return None

        # Isolate unique target files to avoid duplicates bloating the join
        df = df[["FILE_NO"]].drop_duplicates()
        df["FILE_NO"] = df["FILE_NO"].astype(str).str.strip()
        return df

    except Exception as e:
        st.error(f"Error reading cancellation file: {e}")
        return None

# -------------------- SIDEBAR FILE UPLOADERS --------------------
st.sidebar.header("📁 Branch File Uploaders")
cck_file = st.sidebar.file_uploader("Upload CCK Branch Excel", type=["xlsx", "xls"])
lst_file = st.sidebar.file_uploader("Upload LST Branch Excel", type=["xlsx", "xls"])
tlt_file = st.sidebar.file_uploader("Upload TLT Branch Excel", type=["xlsx", "xls"])

st.sidebar.write("---")
st.sidebar.header("🛑 Operational Deductions")
cxl_file = st.sidebar.file_uploader("Upload Sales Cancellation Excel", type=["xlsx", "xls"])

# -------------------- PROCESS BRANCH FILES --------------------
branch_dfs = {}
if cck_file:
    branch_dfs["CCK"] = process_branch_file(cck_file, "CCK")
if lst_file:
    branch_dfs["LST"] = process_branch_file(lst_file, "LST")
if tlt_file:
    branch_dfs["TLT"] = process_branch_file(tlt_file, "TLT")

# Filter out any None or empty DataFrames
active_branches = {k: v for k, v in branch_dfs.items() if v is not None and not v.empty}

if not active_branches:
    st.info("👋 Welcome! Please upload at least one valid Branch Excel data sheet via the sidebar to initialise the dashboard panels.")
    st.stop()

# -------------------- CONSOLIDATE DATA --------------------
total_df = pd.concat(active_branches.values(), ignore_index=True)

# -------------------- APPLY CANCELLATIONS --------------------
cxl_count = 0
if cxl_file:
    with st.spinner("Processing cancellations..."):
        df_cxl = process_cancellation_file(cxl_file)
        if df_cxl is not None and not df_cxl.empty:
            total_df["FILE_NO"] = total_df["FILE_NO"].astype(str).str.strip()

            # Left join exclusively on FILE_NO
            merged = total_df.merge(
                df_cxl,
                on="FILE_NO",
                how="left",
                indicator=True
            )

            cxl_mask = merged["_merge"] == "both"
            cxl_count = cxl_mask.sum()

            # Keep only non‑cancelled rows
            total_df = merged[~cxl_mask].drop(columns=["_merge"]).copy()

            # Free memory
            del merged, df_cxl

            # Update branch DataFrames
            for name in list(active_branches.keys()):
                active_branches[name] = total_df[total_df["Branch"] == name].copy()
                if active_branches[name].empty:
                    del active_branches[name]

            # If everything is cancelled, stop
            if total_df.empty:
                st.warning("⚠️ Deductions successfully applied! All data from uploaded branches was cancelled out by the cancellation register records.")
                st.stop()

grand_corporate_revenue = total_df['NETMAINPRODUCT'].sum()

# -------------------- BUILD TABS --------------------
tab_titles = ["Consolidated Overview"] + [f"{name} Branch" for name in active_branches.keys()]
tabs = st.tabs(tab_titles)

# ==================== TAB 0: CONSOLIDATED OVERVIEW ====================
with tabs[0]:
    st.header("🏢 Unified Corporate Overview")

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Net Combined Sales", format_currency(grand_corporate_revenue))
    c2.metric("Active Operating Branches", len(active_branches))
    if cxl_file and cxl_count > 0:
        c3.metric("Net Confirmed Transactions", f"{len(total_df):,}", delta=f"-{cxl_count} Cancelled")
    else:
        c3.metric("Total Confirmed Transactions", f"{len(total_df):,}")

    st.write("---")

    st.subheader("Overall Performance by Agency Hierarchy")
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
    global_pivot["Overall Corporate Contribution %"] = (global_pivot["Total Contribution ($)"] / grand_corporate_revenue) * 100
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
        branch_pivot["Agency Contribution in Branch %"] = (branch_pivot["Total Branch Sales ($)"] / b_total) * 100
        branch_pivot["Agency Branch Contribution in Overall %"] = (branch_pivot["Total Branch Sales ($)"] / grand_corporate_revenue) * 100
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
        display_raw = format_currency_df(display_raw, ["NETMAINPRODUCT"])
        st.dataframe(display_raw, use_container_width=True, hide_index=True)
