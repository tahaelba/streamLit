import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Trello Sprint Dashboard", layout="wide")

st.title("📊 Trello Sprint Dashboard")

# File upload
uploaded_file = st.file_uploader("Upload Trello CSV", type=["csv"])
if uploaded_file:
    df = pd.read_csv(uploaded_file, parse_dates=["date"])

    # Filters
    members = st.multiselect("Filter by member", sorted(df["member"].dropna().unique()))
    if members:
        df = df[df["member"].isin(members)]

    date_range = st.date_input(
        "Filter by date range",
        [df["date"].min(), df["date"].max()]
    )
    if len(date_range) == 2:
        start, end = date_range
        start = pd.to_datetime(start)
        end = pd.to_datetime(end)
    
        # Ensure df["date"] is datetime, coerce errors to NaT
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        df = df.dropna(subset=["date"])
    
        # Now filter safely
        df = df[(df["date"] >= start) & (df["date"] <= end)]


    st.markdown(f"**Total tickets:** {df['card_id'].nunique()}")

    # Tickets created
    created_df = df[df["action_type"] == "createCard"].groupby("member")["card_id"].nunique().reset_index()
    fig_created = px.bar(created_df, x="member", y="card_id", title="Tickets Created by Member", color="member")
    st.plotly_chart(fig_created, use_container_width=True)

    # Tickets treated (moved to Testing/Done)
    treated_df = df[(df["action_type"] == "updateCard") &
                    (df["list_after"].str.lower().isin(["testing", "done", "terminé", "finished"]))] \
                    .groupby("member")["card_id"].nunique().reset_index()
    fig_treated = px.bar(treated_df, x="member", y="card_id", title="Tickets Treated by Member", color="member")
    st.plotly_chart(fig_treated, use_container_width=True)

    # Comments
    comments_df = df[df["action_type"] == "commentCard"].groupby("member")["card_id"].count().reset_index()
    fig_comments = px.bar(comments_df, x="member", y="card_id", title="Comments by Member", color="member")
    st.plotly_chart(fig_comments, use_container_width=True)

    # Most used tags
    label_counts = df["labels"].dropna().str.split(",").explode().value_counts().reset_index()
    label_counts.columns = ["label_id", "count"]
    fig_labels = px.pie(label_counts, names="label_id", values="count", title="Most Used Labels/Tags")
    st.plotly_chart(fig_labels, use_container_width=True)

    # Activity over time
    daily_activity = df.groupby(df["date"].dt.date)["card_id"].nunique().reset_index()
    daily_activity.columns = ["date", "tickets"]
    fig_timeline = px.line(daily_activity, x="date", y="tickets", title="Tickets Activity Over Time")
    st.plotly_chart(fig_timeline, use_container_width=True)
else:
    st.info("Please upload the `trello_dashboard_data.csv` file to view the dashboard.")
