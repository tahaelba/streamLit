
import io
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="Sales Dashboard", layout="wide")

st.title("📊 Sales Strategy & Pipeline Dashboard")

uploaded = st.file_uploader("Upload the refined Excel template", type=["xlsx"])

def coerce_to_percentage(series):
    \"\"\"
    Convert a pandas Series with possible values like '25', '25%', '0.25', or ' 50 % ' into 0-100 float.
    Non-convertible values become NaN.
    \"\"\"
    def parse_val(v):
        if pd.isna(v):
            return float('nan')
        s = str(v).strip()
        # Replace commas, remove spaces
        s = s.replace(",", ".").replace(" ", "")
        # Strip trailing %
        if s.endswith(\"%\"):
            s = s[:-1]
        try:
            num = float(s)
            # If between 0 and 1, assume it's a fraction and scale to 0-100
            if 0 <= num <= 1:
                num = num * 100
            # Clamp
            num = max(0, min(100, num))
            return num
        except:
            return float('nan')
    return series.apply(parse_val)

if uploaded is None:
    st.info("Please upload the **Refined_Sales_Template.xlsx** (or your filled monthly file) to begin.")
    st.stop()

# Try loading sheets safely
required_sheets = ["Month Strategy", "Companies", "Reservations", "Meetings"]
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

tab1, tab2, tab3, tab4 = st.tabs([\"🧩 Workstreams\", \"🏢 Companies\", \"🏨 Reservations\", \"📅 Meetings\"])

# ---------- TAB 1: Month Strategy (Workstreams) ----------
with tab1:
    st.subheader(\"Workstreams — progress (% status)\")

    ms = pd.read_excel(uploaded, sheet_name=\"Month Strategy\")
    # Normalize column names
    ms.columns = [c.strip() for c in ms.columns]
    expected_cols = [\"Workstream\", \"Status\", \"Comments\", \"Deadline\"]
    for col in expected_cols:
        if col not in ms.columns:
            st.error(f\"`Month Strategy` sheet must contain column: {col}\")
            st.stop()

    # Ensure Status is percentage 0-100
    ms[\"Status_%\"] = coerce_to_percentage(ms[\"Status\"])

    # Warn on rows that couldn't be parsed
    bad_rows = ms[ms[\"Status_%\"].isna() & ms[\"Status\"].notna()]
    if not bad_rows.empty:
        with st.expander(\"⚠️ Some Status values could not be parsed. Click to review.\"):
            st.dataframe(bad_rows[[\"Workstream\", \"Status\"]])

    # Display progress per workstream
    left, right = st.columns([1, 2])
    with left:
        st.write(\"**Progress per Workstream**\")
        for _, r in ms.iterrows():
            ws = str(r[\"Workstream\"])
            val = r[\"Status_%\"]
            if pd.isna(val):
                st.text(f\"• {ws}: (no % provided)\")
            else:
                st.text(f\"{ws}: {val:.0f}%\")
                st.progress(int(val) / 100.0)

    with right:
        st.write(\"**Workstream Progress (Bar Chart)**\")
        plot_df = ms.dropna(subset=[\"Status_%\", \"Workstream\"])
        if plot_df.empty:
            st.info(\"No numeric percentage values found to plot.\")
        else:
            fig = px.bar(plot_df, x=\"Workstream\", y=\"Status_%\", hover_data=[\"Comments\", \"Deadline\"], title=\"Workstream % Completion\")
            fig.update_yaxes(range=[0, 100], title=\"%\")
            st.plotly_chart(fig, use_container_width=True)

    # Sorting option
    sort_order = st.radio(\"Sort by\", [\"Workstream (A-Z)\", \"Status (High → Low)\", \"Status (Low → High)\"], horizontal=True)
    if sort_order == \"Workstream (A-Z)\":
        view = ms.sort_values(\"Workstream\", kind=\"stable\")
    elif sort_order == \"Status (High → Low)\":
        view = ms.sort_values(\"Status_%\", ascending=False, kind=\"stable\")
    else:
        view = ms.sort_values(\"Status_%\", ascending=True, kind=\"stable\")
    st.dataframe(view[[\"Workstream\", \"Status\", \"Status_%\", \"Comments\", \"Deadline\"]].rename(columns={\"Status_%\":\"Status (%)\"}))

# ---------- TAB 2: Companies ----------
with tab2:
    st.subheader(\"Company Pipeline\")
    comp = pd.read_excel(uploaded, sheet_name=\"Companies\")
    comp.columns = [c.strip() for c in comp.columns]
    # Flexible columns (template recommended fields)
    # Required: Company. Optional: Pricing, Contract 2026, Payment, Status, Comments
    if \"Company\" not in comp.columns:
        st.error(\"`Companies` sheet must contain a 'Company' column.\")
        st.stop()

    # Status ordering (based on text). Users can customize mapping if needed.
    # We'll infer simple stages from Status or Comments if present.
    stage_map = {
        \"lead\": 1,
        \"contacted\": 2,
        \"meeting\": 3,
        \"proposal\": 4,
        \"negotiation\": 5,
        \"waiting\": 6,
        \"contract\": 7,
        \"won\": 8,
        \"lost\": 0
    }
    def infer_stage(row):
        text = \"\"
        for col in [\"Status\", \"Comments\"]:
            if col in comp.columns and pd.notna(row.get(col)):
                text += \" \" + str(row.get(col)).lower()
        score = 0
        for k, v in stage_map.items():
            if k in text:
                score = max(score, v)
        return score

    comp[\"StageScore\"] = comp.apply(infer_stage, axis=1)

    # Filters
    with st.sidebar:
        st.header(\"Filters\")
        companies = sorted(comp[\"Company\"].dropna().unique().tolist())
        selected_companies = st.multiselect(\"Companies\", companies, default=companies)

    comp_f = comp[comp[\"Company\"].isin(selected_companies)]

    # Chart: Companies by stage
    by_stage = comp_f.groupby(\"StageScore\").size().reset_index(name=\"Count\")
    if not by_stage.empty:
        fig_s = px.bar(by_stage, x=\"StageScore\", y=\"Count\", title=\"Companies by Stage (inferred)\")
        st.plotly_chart(fig_s, use_container_width=True)

    st.write(\"**Companies (ordered by inferred stage)**\")
    st.dataframe(comp_f.sort_values([\"StageScore\", \"Company\"], ascending=[False, True], kind=\"stable\"))

# ---------- TAB 3: Reservations ----------
with tab3:
    st.subheader(\"Reservations by Company & City\")
    res = pd.read_excel(uploaded, sheet_name=\"Reservations\")
    res.columns = [c.strip() for c in res.columns]
    required_cols = [\"Company\", \"Nights\", \"Amount (MAD)\", \"City\"]
    for col in required_cols:
        if col not in res.columns:
            st.error(f\"`Reservations` sheet must contain column: {col}\")
            st.stop()

    res[\"Nights\"] = pd.to_numeric(res[\"Nights\"], errors=\"coerce\")
    res[\"Amount (MAD)\"] = pd.to_numeric(res[\"Amount (MAD)\"], errors=\"coerce\")

    # Filters
    companies_r = sorted(res[\"Company\"].dropna().unique().tolist())
    cities = sorted(res[\"City\"].dropna().unique().tolist())
    c1, c2 = st.columns(2)
    with c1:
        selected_companies_r = st.multiselect(\"Companies\", companies_r, default=companies_r)
    with c2:
        selected_cities = st.multiselect(\"Cities\", cities, default=cities)

    res_f = res[res[\"Company\"].isin(selected_companies_r) & res[\"City\"].isin(selected_cities)]

    # Chart: Amount by Company & City
    if not res_f.empty:
        fig_r = px.bar(res_f, x=\"Company\", y=\"Amount (MAD)\", color=\"City\", barmode=\"group\", title=\"Revenue by Company & City\")
        st.plotly_chart(fig_r, use_container_width=True)

    # Pivot: Nights by City
    pivot = res_f.pivot_table(index=\"Company\", columns=\"City\", values=\"Nights\", aggfunc=\"sum\").fillna(0)
    st.write(\"**Nights by City (pivot)**\")
    st.dataframe(pivot)

# ---------- TAB 4: Meetings ----------
with tab4:
    st.subheader(\"Meetings Summary\")
    meet = pd.read_excel(uploaded, sheet_name=\"Meetings\")
    meet.columns = [c.strip() for c in meet.columns]

    # Flexible schema: require at least Company and Date if present
    if \"Company\" not in meet.columns:
        st.error(\"`Meetings` sheet must contain a 'Company' column.\")
        st.stop()

    # Parse a DateTime if present
    if \"Date\" in meet.columns:
        meet[\"Date\"] = pd.to_datetime(meet[\"Date\"], errors=\"coerce\", dayfirst=True)

    # Filters
    companies_m = sorted(meet[\"Company\"].dropna().unique().tolist())
    selected_companies_m = st.multiselect(\"Companies\", companies_m, default=companies_m)

    meet_f = meet[meet[\"Company\"].isin(selected_companies_m)]

    # Chart: Meetings per company
    by_company = meet_f.groupby(\"Company\").size().reset_index(name=\"Meetings\")
    if not by_company.empty:
        fig_m = px.bar(by_company, x=\"Company\", y=\"Meetings\", title=\"Meetings per Company\")
        st.plotly_chart(fig_m, use_container_width=True)

    st.write(\"**Meetings (table)**\")
    st.dataframe(meet_f)
