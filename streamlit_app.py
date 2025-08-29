import io
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="Sales Dashboard", layout="wide")

st.title("📊 Sales Strategy & Pipeline Dashboard")

uploaded = st.file_uploader("Upload the refined Excel template", type=["xlsx"])

def coerce_to_percentage(series):
    """
    Convert values like '25', '25%', '0.25', or ' 50 % ' into a 0–100 float.
    Non-convertible values become NaN.
    """
    def parse_val(v):
        if pd.isna(v):
            return float('nan')
        s = str(v).strip()
        s = s.replace(",", ".").replace(" ", "")
        if s.endswith("%"):
            s = s[:-1]
        try:
            num = float(s)
            if 0 <= num <= 1:
                num = num * 100
            num = max(0, min(100, num))
            return num
        except Exception:
            return float('nan')
    return series.apply(parse_val)

if uploaded is None:
    st.info("Please upload the **Refined_Sales_Template.xlsx** (or your filled monthly file) to begin.")
    st.stop()

# Load sheets safely
required_sheets = ["Month Strategy", "Reservations", "Meetings"]
try:
    xls = pd.ExcelFile(uploaded)
    available = xls.sheet_names
    missing = [s for s in required_sheets if s not in available]
    if missing:
        st.error(f"Missing sheets: {missing}. Please use the refined template structure.")
        st.stop()
except Exception as e:
    st.error(f"Failed to read Excel: {e}")
    st.stop()

tab1, tab3, tab4 = st.tabs(["🧩 Workstreams", "🏨 Reservations", "📅 Meetings"])

# ---------- TAB 1: Month Strategy ----------
with tab1:
    st.subheader("Workstreams — progress (% status)")
    ms = pd.read_excel(uploaded, sheet_name="Month Strategy")
    ms.columns = [c.strip() for c in ms.columns]

    expected_cols = ["Workstream", "Status", "Comments", "Deadline"]
    for col in expected_cols:
        if col not in ms.columns:
            st.error(f"`Month Strategy` sheet must contain column: {col}")
            st.stop()

    ms["Status_%"] = coerce_to_percentage(ms["Status"])

    left, right = st.columns([1, 2])
    with left:
        st.write("**Progress per Workstream**")
        for _, r in ms.iterrows():
            ws = str(r["Workstream"])
            val = r["Status_%"]
            if pd.isna(val):
                st.text(f"• {ws}: (no % provided)")
            else:
                st.text(f"{ws}: {val:.0f}%")
                st.progress(int(val) / 100.0)
    with right:
        plot_df = ms.dropna(subset=["Status_%", "Workstream"])
        if not plot_df.empty:
            fig = px.bar(
                plot_df,
                x="Workstream",
                y="Status_%",
                hover_data=["Comments", "Deadline"],
                title="Workstream % Completion"
            )
            fig.update_yaxes(range=[0, 100], title="%")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No numeric percentage values found to plot.")

    # Optional sortable table
    sort_order = st.radio("Sort by", ["Workstream (A–Z)", "Status (High → Low)", "Status (Low → High)"], horizontal=True)
    if sort_order == "Workstream (A–Z)":
        view = ms.sort_values("Workstream", kind="stable")
    elif sort_order == "Status (High → Low)":
        view = ms.sort_values("Status_%", ascending=False, kind="stable")
    else:
        view = ms.sort_values("Status_%", ascending=True, kind="stable")
    st.dataframe(
        view[["Workstream", "Status_%", "Comments", "Deadline"]]
        .rename(columns={"Status_%": "Progress (%)"})
    )

# # ---------- TAB 2: Companies ----------
# with tab2:
#     st.subheader("Company Pipeline")
#     comp = pd.read_excel(uploaded, sheet_name="Companies")
#     comp.columns = [c.strip() for c in comp.columns]

#     if "Company" not in comp.columns:
#         st.error("`Companies` sheet must contain a 'Company' column.")
#         st.stop()

#     # Show labels (not scores). Keep list to preserve an intended order if needed.
#     stage_map = [
#         "lead",
#         "contacted",
#         "meeting",
#         "proposal",
#         "negotiation",
#         "waiting",
#         "contract",
#         "won",
#         "lost",
#     ]

#     def infer_stage(row):
#         text = ""
#         for col in ["Status", "Comments"]:
#             if col in comp.columns and pd.notna(row.get(col)):
#                 text += " " + str(row.get(col)).lower()

#         # Explicit rule: "contract" OR "negotiation" => Negotiation (unified)
#         if "contract" in text or "negotiation" in text:
#             return "Negotiation"

#         for s in stage_map:
#             if s in text:
#                 return s.capitalize()
#         return "Unspecified"

#     comp["Stage"] = comp.apply(infer_stage, axis=1)

#     # Sidebar filters
#     with st.sidebar:
#         st.header("Filters")
#         companies = sorted(comp["Company"].dropna().unique().tolist())
#         selected_companies = st.multiselect("Companies", companies, default=companies)
#         stages = sorted(comp["Stage"].dropna().unique().tolist())
#         selected_stages = st.multiselect("Stages", stages, default=stages)

#     comp_filtered = comp[comp["Company"].isin(selected_companies) & comp["Stage"].isin(selected_stages)]

#     # Chart: Companies by Stage (labels)
#     by_stage = comp_filtered.groupby("Stage").size().reset_index(name="Count")
#     if not by_stage.empty:
#         fig_s = px.bar(by_stage, x="Stage", y="Count", title="Companies by Stage")
#         st.plotly_chart(fig_s, use_container_width=True)
#     else:
#         st.info("No companies match the selected filters.")

#     st.write("**Companies (with inferred stage)**")
#     st.dataframe(comp_filtered.sort_values(["Stage", "Company"], kind="stable"))

# ---------- TAB 3: Reservations ----------
with tab3:
    st.subheader("Reservations by Company & City")
    res = pd.read_excel(uploaded, sheet_name="Reservations")
    res.columns = [c.strip() for c in res.columns]

    required_cols = ["Company", "Nights", "Amount (MAD)", "City"]
    for col in required_cols:
        if col not in res.columns:
            st.error(f"`Reservations` sheet must contain column: {col}")
            st.stop()

    res["Nights"] = pd.to_numeric(res["Nights"], errors="coerce")
    res["Amount (MAD)"] = pd.to_numeric(res["Amount (MAD)"], errors="coerce")

    companies_r = sorted(res["Company"].dropna().unique().tolist())
    cities = sorted(res["City"].dropna().unique().tolist())
    c1, c2 = st.columns(2)
    with c1:
        selected_companies_r = st.multiselect("Companies", companies_r, default=companies_r)
    with c2:
        selected_cities = st.multiselect("Cities", cities, default=cities)

    res_f = res[res["Company"].isin(selected_companies_r) & res["City"].isin(selected_cities)]

    if not res_f.empty:
        fig_r = px.bar(
            res_f, x="Company", y="Amount (MAD)", color="City",
            barmode="group", title="Revenue by Company & City"
        )
        st.plotly_chart(fig_r, use_container_width=True)
    else:
        st.info("No reservations match the selected filters.")

    pivot = res_f.pivot_table(index="Company", columns="City", values="Nights", aggfunc="sum").fillna(0)
    st.write("**Nights by City (pivot)**")
    st.dataframe(pivot)

# ---------- TAB 4: Meetings ----------
with tab4:
    st.subheader("Meetings Summary")
    meet = pd.read_excel(uploaded, sheet_name="Meetings")
    meet.columns = [c.strip() for c in meet.columns]

    if "Company" not in meet.columns:
        st.error("`Meetings` sheet must contain a 'Company' column.")
        st.stop()

    # Optional parsing
    if "Date" in meet.columns:
        meet["Date"] = pd.to_datetime(meet["Date"], errors="coerce", dayfirst=True)

    # Normalize Status values (supports done/pending/confirmed in any case)
    if "Status" in meet.columns:
        meet["Status"] = meet["Status"].fillna("Unspecified").astype(str).str.strip().str.lower()
        status_map = {
            "done": "Done",
            "completed": "Done",
            "confirm": "Confirmed",
            "confirmed": "Confirmed",
            "pending": "Pending",
            "unspecified": "Unspecified",
            "": "Unspecified",
        }
        meet["Status"] = meet["Status"].map(status_map).fillna("Other")
    else:
        meet["Status"] = "Unspecified"

    # Filters
    c1, c2 = st.columns(2)
    with c1:
        companies_m = sorted(meet["Company"].dropna().unique().tolist())
        selected_companies_m = st.multiselect("Companies", companies_m, default=companies_m)
    with c2:
        statuses_m = sorted(meet["Status"].dropna().unique().tolist())
        selected_statuses_m = st.multiselect("Statuses", statuses_m, default=statuses_m)

    meet_f = meet[meet["Company"].isin(selected_companies_m) & meet["Status"].isin(selected_statuses_m)]

    # Chart 1: Meetings per company (overall count)
    by_company = meet_f.groupby("Company").size().reset_index(name="Meetings")
    if not by_company.empty:
        fig_m = px.bar(by_company, x="Company", y="Meetings", title="Meetings per Company")
        st.plotly_chart(fig_m, use_container_width=True)

    # Chart 2: Stacked by Status
    by_company_status = meet_f.groupby(["Company", "Status"]).size().reset_index(name="Count")
    if not by_company_status.empty:
        fig_status = px.bar(
            by_company_status,
            x="Company", y="Count", color="Status",
            barmode="stack", title="Meetings per Company by Status"
        )
        st.plotly_chart(fig_status, use_container_width=True)

    # Chart 3: Overall status distribution
    by_status = meet_f["Status"].value_counts().reset_index()
    if not by_status.empty:
        by_status.columns = ["Status", "Count"]
        fig_pie = px.pie(by_status, names="Status", values="Count", title="Overall Meeting Status")
        st.plotly_chart(fig_pie, use_container_width=True)

    # Table
    cols_show = [c for c in ["Company", "Contact Person", "Date", "Time", "Mode", "Status", "Notes"] if c in meet_f.columns]
    st.write("**Meetings (table)**")
    st.dataframe(meet_f[cols_show] if cols_show else meet_f)
