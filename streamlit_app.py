
import io
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="Sales Dashboard", layout="wide")

st.title("📊 Sales Strategy & Pipeline Dashboard")

uploaded = st.file_uploader("Upload the refined Excel template", type=["xlsx"])

def coerce_to_percentage(series):
    """Convert values like '25', '25%', '0.25' into 0–100 float"""
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
        except:
            return float('nan')
    return series.apply(parse_val)

if uploaded is None:
    st.info("Please upload the **Refined_Sales_Template.xlsx** (or your filled monthly file) to begin.")
    st.stop()

# Load sheets
required_sheets = ["Month Strategy", "Companies", "Reservations"]
xls = pd.ExcelFile(uploaded)
available = xls.sheet_names
missing = [s for s in required_sheets if s not in available]
if missing:
    st.error(f"Missing sheets: {missing}. Please use the refined template structure.")
    st.stop()

tab1, tab2, tab3, tab4 = st.tabs(["🧩 Workstreams", "🏢 Companies", "🏨 Reservations"])

# ---------- TAB 1: Month Strategy ----------
with tab1:
    st.subheader("Workstreams — progress (% status)")
    ms = pd.read_excel(uploaded, sheet_name="Month Strategy")
    ms.columns = [c.strip() for c in ms.columns]
    ms["Status_%"] = coerce_to_percentage(ms["Status"])

    left, right = st.columns([1, 2])
    with left:
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
            fig = px.bar(plot_df, x="Workstream", y="Status_%", hover_data=["Comments", "Deadline"], title="Workstream % Completion")
            fig.update_yaxes(range=[0, 100], title="%")
            st.plotly_chart(fig, use_container_width=True)

# ---------- TAB 2: Companies ----------
with tab2:
    st.subheader("Company Pipeline")
    comp = pd.read_excel(uploaded, sheet_name="Companies")
    comp.columns = [c.strip() for c in comp.columns]

    if "Company" not in comp.columns:
        st.error("`Companies` sheet must contain a 'Company' column.")
        st.stop()

    # Stage labels instead of scores
    stage_map = [
        "lead",
        "contacted",
        "meeting",
        "proposal",
        "negotiation",
        "waiting",
        "contract",
        "won",
        "lost"
    ]

    def infer_stage(row):
        text = ""
        for col in ["Status", "Comments"]:
            if col in comp.columns and pd.notna(row.get(col)):
                text += " " + str(row.get(col)).lower()
        for s in stage_map:
            if s in text:
                return s.capitalize()
        return "Unspecified"

    comp["Stage"] = comp.apply(infer_stage, axis=1)

    # Chart: Companies by Stage
    by_stage = comp.groupby("Stage").size().reset_index(name="Count")
    if not by_stage.empty:
        fig_s = px.bar(by_stage, x="Stage", y="Count", title="Companies by Stage")
        st.plotly_chart(fig_s, use_container_width=True)

    st.write("**Companies (with inferred stage)**")
    st.dataframe(comp.sort_values(["Stage", "Company"], kind="stable"))

# ---------- TAB 3: Reservations ----------
with tab3:
    st.subheader("Reservations by Company & City")
    res = pd.read_excel(uploaded, sheet_name="Reservations")
    res.columns = [c.strip() for c in res.columns]
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
        fig_r = px.bar(res_f, x="Company", y="Amount (MAD)", color="City", barmode="group", title="Revenue by Company & City")
        st.plotly_chart(fig_r, use_container_width=True)

    pivot = res_f.pivot_table(index="Company", columns="City", values="Nights", aggfunc="sum").fillna(0)
    st.dataframe(pivot)

# ---------- TAB 4: Meetings ----------
# with tab4:
#     st.subheader("Meetings Summary")
#     meet = pd.read_excel(uploaded, sheet_name="Meetings")
#     meet.columns = [c.strip() for c in meet.columns]
#     if "Date" in meet.columns:
#         meet["Date"] = pd.to_datetime(meet["Date"], errors="coerce", dayfirst=True)

#     companies_m = sorted(meet["Company"].dropna().unique().tolist())
#     selected_companies_m = st.multiselect("Companies", companies_m, default=companies_m)
#     meet_f = meet[meet["Company"].isin(selected_companies_m)]

#     by_company = meet_f.groupby("Company").size().reset_index(name="Meetings")
#     if not by_company.empty:
#         fig_m = px.bar(by_company, x="Company", y="Meetings", title="Meetings per Company")
#         st.plotly_chart(fig_m, use_container_width=True)

#     st.dataframe(meet_f)
