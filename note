"""
Bank Submit History — Streamlit Dashboard
==========================================
File: bank_dashboard_generator.py

Setup (once):
    pip install streamlit pandas plotly openpyxl

Run:
    streamlit run bank_dashboard_generator.py
"""

import os, glob, shutil, tempfile
from datetime import datetime
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from io import BytesIO

try:
    from reportlab.lib.pagesizes import A3, landscape
    from reportlab.lib import colors
    from reportlab.lib.units import inch
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.enums import TA_LEFT
    from reportlab.platypus import SimpleDocTemplate, LongTable, TableStyle, Paragraph, PageBreak
    from reportlab.pdfbase import pdfmetrics
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Bank Submit Dashboard", page_icon="🏦",
                   layout="wide", initial_sidebar_state="expanded")


PL = dict(
    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
    font=dict(color="#8899aa", size=11, family="monospace"),
    margin=dict(l=8, r=8, t=32, b=8),
    legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color="#8899aa", size=10)),
    xaxis=dict(gridcolor="#1a2a3a", linecolor="#1a2a3a"),
    yaxis=dict(gridcolor="#1a2a3a", linecolor="#1a2a3a"),
)
C = ["#00c9a7","#1a8fff","#ff6b35","#ffd700","#cc44ff",
     "#ff3366","#44aaff","#ff9944","#66dd66","#00aaff"]

REQUIRED_COLUMNS = [
    "Firm Name", "Sales Person", "Bank Submition Date", "Invoice Value",
    "Lc Value", "Maturity Date", "Payment. Rcv Dt", "Bank Accept Date",
    "LC No", "Our Bank", "Party Name", "Bank Name"
]

PL_GENERAL = {k: v for k, v in PL.items() if k not in ("legend", "xaxis", "yaxis")}

def usd(v):
    try:
        v = float(v)
    except Exception:
        return "$0.00"
    if v >= 1e6: return f"${v/1e6:.2f}M"
    if v >= 1e3: return f"${v/1e3:.1f}K"
    return f"${v:.2f}"


def sh(label):
    st.markdown(f'<p class="sh">{label}</p>', unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADER
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="⏳ Loading data…")
def load(path):
    try:
        df = pd.read_excel(path, sheet_name="Raw Data")
    except Exception:
        df = pd.read_excel(path)
    df.columns = df.columns.str.strip()
    # Drop stray unnamed index columns and empty columns from Excel imports
    df = df.loc[:, ~df.columns.str.match(r"^Unnamed")]  # remove Unnamed: 0 / index cols
    df = df.dropna(axis=1, how="all")
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            "Excel file is missing required columns: " + ", ".join(missing)
        )
    df = df.dropna(subset=["Firm Name"])
    df["Sales Person"] = (df["Sales Person"].astype(str).str.strip()
                          .str.replace("_x000D_\n", "", regex=False).str.strip())
    df["Sales Person"] = df["Sales Person"].replace(
        r"^(nan|NaN|none|None|\s*)$", None, regex=True
    )
    df.loc[df["Sales Person"].isin([None, ""]), "Sales Person"] = None
    df["_date"]     = pd.to_datetime(df["Bank Submition Date"], errors="coerce")
    df["MonthSort"] = df["_date"].dt.to_period("M")
    df["Month"]     = df["_date"].dt.strftime("%b %Y")
    df["WeekSort"]  = df["_date"].dt.to_period("W")
    df["Week"]      = df["_date"].dt.strftime("W%V") + " " + df["_date"].dt.strftime("%b %Y")
    df["DayName"]   = df["_date"].dt.strftime("%a")
    df["Date"]      = df["_date"].dt.strftime("%d %b %Y")
    return df

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
st.sidebar.markdown("## 🏦 Bank Submit\nDashboard")
st.sidebar.markdown("---")

xlsx = [f for f in glob.glob("*.xlsx") + glob.glob("**/*.xlsx", recursive=True)
        if "Dashboard" not in f]
up = st.sidebar.file_uploader("📂 Upload Excel file", type=["xlsx"])
if up:
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    shutil.copyfileobj(up, tmp); tmp.close(); FP = tmp.name
elif xlsx:
    FP = st.sidebar.selectbox("Or pick a file", xlsx)
else:
    st.warning("⚠️ Please upload your Excel file using the sidebar."); st.stop()

raw = load(FP)

st.sidebar.markdown("### 🔽 Filters")
ml  = [str(m) for m in sorted(raw["MonthSort"].dropna().unique())]
sm  = st.sidebar.multiselect("Month", ml, default=ml)
sf  = st.sidebar.multiselect("Firm", sorted(raw["Firm Name"].dropna().unique()), default=sorted(raw["Firm Name"].dropna().unique()))
sb  = st.sidebar.multiselect("Our Bank", sorted(raw["Our Bank"].dropna().unique()), default=sorted(raw["Our Bank"].dropna().unique()))

sales_persons = sorted(raw["Sales Person"].dropna().unique())
ss_choices = ["(Blank)"] + sales_persons if raw["Sales Person"].isna().any() else sales_persons
ss = st.sidebar.multiselect("Sales Person", ss_choices, default=ss_choices)

sparty = st.sidebar.multiselect("Party Name", sorted(raw["Party Name"].dropna().unique()), default=sorted(raw["Party Name"].dropna().unique()))

min_date = raw["_date"].min()
max_date = raw["_date"].max()
if pd.isna(min_date) or pd.isna(max_date):
    date_range = st.sidebar.date_input("Date Range", value=(pd.Timestamp.today().date(), pd.Timestamp.today().date()))
else:
    date_range = st.sidebar.date_input("Date Range", value=(min_date.date(), max_date.date()))

df = raw.copy()
if sm:
    df = df[df["MonthSort"].astype(str).isin(sm)]
if sf:
    df = df[df["Firm Name"].isin(sf)]
if sb:
    df = df[df["Our Bank"].isin(sb)]
if ss:
    if "(Blank)" in ss:
        selected = [s for s in ss if s != "(Blank)"]
        df = df[(df["Sales Person"].isin(selected)) | df["Sales Person"].isna()]
    else:
        df = df[df["Sales Person"].isin(ss)]
if sparty:
    df = df[df["Party Name"].isin(sparty)]
if isinstance(date_range, tuple) and len(date_range) == 2:
    start_date, end_date = date_range
    df = df[(df["_date"].dt.date >= start_date) & (df["_date"].dt.date <= end_date)]

st.sidebar.markdown("---")
st.sidebar.caption(f"Showing **{len(df):,}** of **{len(raw):,}** records")
if df.empty: st.warning("No records match the filters."); st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# AGGREGATES — MONTHLY
# ─────────────────────────────────────────────────────────────────────────────
N          = len(df)
inv        = df["Invoice Value"].sum()
lc         = df["Lc Value"].sum()
mat_v      = df[df["Maturity Date"].notna()]["Invoice Value"].sum()
pay_v      = df[df["Payment. Rcv Dt"].notna()]["Invoice Value"].sum()
paid_n     = int(df["Payment. Rcv Dt"].notna().sum())
acc_n      = int((df["Bank Accept Date"].notna() & df["Payment. Rcv Dt"].isna()).sum())
nacc_n     = int(df["Bank Accept Date"].isna().sum())
acc_v      = df[df["Bank Accept Date"].notna() & df["Payment. Rcv Dt"].isna()]["Invoice Value"].sum()
nacc_v     = df[df["Bank Accept Date"].isna()]["Invoice Value"].sum()

monthly    = (df.groupby(["MonthSort","Month"])
               .agg(Count=("LC No","count"), Inv=("Invoice Value","sum"), LC=("Lc Value","sum"))
               .reset_index().sort_values("MonthSort"))
by_firm    = df.groupby("Firm Name").agg(Inv=("Invoice Value","sum"), N=("LC No","count")).reset_index().sort_values("Inv", ascending=False)
by_bank    = df.groupby("Our Bank").agg(Inv=("Invoice Value","sum"), N=("LC No","count")).reset_index().sort_values("Inv", ascending=False)
t_party    = df.groupby("Party Name").agg(Inv=("Invoice Value","sum"), N=("LC No","count")).reset_index().sort_values("Inv", ascending=False).head(10)
t_bname    = df.groupby("Bank Name").agg(Inv=("Invoice Value","sum"), N=("LC No","count")).reset_index().sort_values("Inv", ascending=False).head(10)

spg        = df[df["Sales Person"].notna()].groupby("Sales Person").agg(Inv=("Invoice Value","sum"), N=("LC No","count")).reset_index().sort_values("Inv", ascending=False)
sp_p       = df[df["Payment. Rcv Dt"].notna() & df["Sales Person"].notna()].groupby("Sales Person").size().reset_index(name="Paid")
spg        = spg.merge(sp_p, on="Sales Person", how="left").fillna(0)
spg["Pct"] = (spg["Paid"] / spg["N"] * 100).round(1)

# ─────────────────────────────────────────────────────────────────────────────
# AGGREGATES — WEEKLY
# ─────────────────────────────────────────────────────────────────────────────
weekly = (df.groupby(["WeekSort","Week"])
           .agg(Count=("LC No","count"), Inv=("Invoice Value","sum"), LC=("Lc Value","sum"),
                Paid_n=("Payment. Rcv Dt", lambda x: x.notna().sum()))
           .reset_index().sort_values("WeekSort"))
weekly["Paid_pct"] = (weekly["Paid_n"] / weekly["Count"] * 100).round(1)

# Weekly by firm (stacked)
wk_firm = (df.groupby(["WeekSort","Week","Firm Name"])
             .agg(Inv=("Invoice Value","sum"), Count=("LC No","count"))
             .reset_index().sort_values("WeekSort"))

# Weekly by sales person
wk_sp = (df[df["Sales Person"].notna()]
           .groupby(["WeekSort","Week","Sales Person"])
           .agg(Inv=("Invoice Value","sum"), Count=("LC No","count"))
           .reset_index().sort_values("WeekSort"))

# Weekly by our bank
wk_bank = (df.groupby(["WeekSort","Week","Our Bank"])
             .agg(Inv=("Invoice Value","sum"), Count=("LC No","count"))
             .reset_index().sort_values("WeekSort"))

# Weekly payment status
wk_status = df.copy()
wk_status["Status"] = wk_status.apply(lambda r:
    "Paid"         if pd.notna(r["Payment. Rcv Dt"]) else
    "Accepted"     if pd.notna(r["Bank Accept Date"]) else
    "Not Accepted", axis=1)
wk_st_grp = (wk_status.groupby(["WeekSort","Week","Status"])
              .agg(Count=("LC No","count"), Inv=("Invoice Value","sum"))
              .reset_index().sort_values("WeekSort"))

# Top party weekly
wk_party = (df.groupby(["WeekSort","Week","Party Name"])
              .agg(Inv=("Invoice Value","sum"), Count=("LC No","count"))
              .reset_index().sort_values("WeekSort"))
top_parties_list = t_party["Party Name"].tolist()
wk_party_top = wk_party[wk_party["Party Name"].isin(top_parties_list)]

period = f"{monthly['Month'].iloc[0]} – {monthly['Month'].iloc[-1]}" if len(monthly) else "—"

# ─────────────────────────────────────────────────────────────────────────────
# HEADER + KPIs
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("# 🏦 Bank Submit History Dashboard")
st.markdown(f"<p style='color:#556677;font-size:12px;margin-top:-12px;'>"
            f"Period: <b style='color:#00c9a7'>{period}</b> &nbsp;|&nbsp; "
            f"File: {os.path.basename(FP)}</p>", unsafe_allow_html=True)
st.markdown("---")

c1,c2,c3,c4 = st.columns(4)
c1.metric("📋 Total Submissions",       f"{N:,}")
c2.metric("💵 Total Invoice Value",     usd(inv))
c3.metric("📅 Maturity Received Value", usd(mat_v))
c4.metric("✅ Payment Received Value",  usd(pay_v),
          delta=f"{paid_n} records · {paid_n/N*100:.1f}%")
st.markdown("")

# ─────────────────────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────────────────────
t1, t2, t3, t4, t5, t6 = st.tabs([
    "📊 Overview",
    "📅 Weekly Analysis",
    "🏢 Firm & Sales Person",
    "🏦 Banks",
    "👥 Top Parties",
    "🔄 Payment Status",
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
with t1:
    l, r = st.columns(2)

    with l:
        sh("📅 Monthly Submission Trend")
        fig = go.Figure()
        fig.add_bar(x=monthly["Month"], y=monthly["Count"], name="Submissions",
                    marker_color="#1a8fff", yaxis="y1")
        fig.add_scatter(x=monthly["Month"], y=monthly["Inv"], name="Invoice Value",
                        mode="lines+markers", line=dict(color="#00c9a7", width=2.5),
                        marker=dict(size=6), yaxis="y2")
        # একবার PL_general বানান
        PL_general = {k: v for k, v in PL.items() if k not in ("legend", "xaxis", "yaxis")}


        # তারপর update_layout এ ব্যবহার করুন
        fig.update_layout(
            **PL_general,
            yaxis=dict(title="Submissions", gridcolor="#1a2a3a"),
            legend=dict(orientation="h", x=0, y=1.1),
            height=340
        )

        st.plotly_chart(fig, use_container_width=True)

    with r:
        sh("🏦 Our Bank — Invoice Value")
        fig2 = px.bar(by_bank, x="Our Bank", y="Inv", color="Our Bank",
                      color_discrete_sequence=C, text=by_bank["Inv"].apply(usd))
        fig2.update_traces(textposition="outside", textfont_size=10)
        fig2.update_layout(**PL, showlegend=False)

        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("---")
    a, b = st.columns(2)
    with a:
        sh("🏢 Firm-wise Invoice Value")
        fig3 = px.pie(by_firm, names="Firm Name", values="Inv",
                      color_discrete_sequence=C, hole=0.45)
        fig3.update_layout(**PL)
        fig3.update_traces(textinfo="label+percent", textfont_size=11)
        st.plotly_chart(fig3, use_container_width=True)
    with b:
        sh("📋 Monthly Summary Table")
        tbl = monthly[["Month","Count","Inv","LC"]].copy()
        tbl.columns = ["Month","Submissions","Invoice Value (USD)","LC Value (USD)"]
        tbl["Invoice Value (USD)"] = tbl["Invoice Value (USD)"].map(lambda x: f"${x:,.2f}")
        tbl["LC Value (USD)"]      = tbl["LC Value (USD)"].map(lambda x: f"${x:,.2f}")
        st.dataframe(tbl, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — WEEKLY ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
with t2:

    # ── Row 1: KPI mini cards ─────────────────────────────────────────────
    best_wk  = weekly.loc[weekly["Count"].idxmax()]
    best_inv = weekly.loc[weekly["Inv"].idxmax()]
    avg_wk   = weekly["Count"].mean()
    avg_inv  = weekly["Inv"].mean()

    w1,w2,w3,w4 = st.columns(4)
    w1.metric("📆 Total Weeks",          f"{len(weekly)}")
    w2.metric("🔝 Best Week (Subs)",     f"{best_wk['Week']} — {int(best_wk['Count'])}")
    w3.metric("💰 Best Week (Value)",    f"{best_inv['Week']} — {usd(best_inv['Inv'])}")
    w4.metric("📊 Weekly Avg Subs",      f"{avg_wk:.0f}")
    st.markdown("")

    # ── Row 2: Submission + Invoice dual axis ─────────────────────────────
    sh("📅 Week-wise Submission Count + Invoice Value")
    fig = go.Figure()
    fig.add_bar(x=weekly["Week"], y=weekly["Count"], name="Submissions",
                marker_color="#1a8fff", yaxis="y1",
                text=weekly["Count"], textposition="outside", textfont=dict(size=9))
    fig.add_scatter(x=weekly["Week"], y=weekly["Inv"], name="Invoice Value (USD)",
                    mode="lines+markers", line=dict(color="#00c9a7", width=2.5),
                    marker=dict(size=5), yaxis="y2")
    fig.update_layout(**PL_general,
        yaxis=dict(title="Submissions", gridcolor="#1a2a3a"),
        yaxis2=dict(title="Invoice Value (USD)", overlaying="y", side="right",
                    gridcolor="rgba(0,0,0,0)", tickformat="$.2s"),
        xaxis=dict(tickangle=-40, tickfont=dict(size=9)),
        legend=dict(orientation="h", x=0, y=1.1),
        height=340)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")

    # ── Row 3: Stacked by Firm + by Our Bank ─────────────────────────────
    l, r = st.columns(2)

    with l:
        sh("🏢 Weekly Invoice Value by Firm (Stacked)")
        fig2 = px.bar(wk_firm, x="Week", y="Inv", color="Firm Name",
                      color_discrete_sequence=C, barmode="stack",
                      text=None)
        fig2.update_layout(**PL_general,
            xaxis=dict(tickangle=-40, tickfont=dict(size=9), title=""),
            yaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            legend=dict(orientation="h", x=0, y=1.1, font=dict(size=9)),
            height=320)
        st.plotly_chart(fig2, use_container_width=True)

    with r:
        sh("🏦 Weekly Invoice Value by Our Bank (Stacked)")
        fig3 = px.bar(wk_bank, x="Week", y="Inv", color="Our Bank",
                      color_discrete_sequence=C, barmode="stack")
        fig3.update_layout(**PL_general,
            xaxis=dict(tickangle=-40, tickfont=dict(size=9), title=""),
            yaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            legend=dict(orientation="h", x=0, y=1.1, font=dict(size=9)),
            height=320)
        st.plotly_chart(fig3, use_container_width=True)

    st.markdown("---")

    # ── Row 4: Sales Person weekly line + Payment status stacked ─────────
    l2, r2 = st.columns(2)

    with l2:
        sh("🧑‍💼 Weekly Invoice by Sales Person (Top 6)")
        top6_sp = spg.head(6)["Sales Person"].tolist()
        wk_sp6  = wk_sp[wk_sp["Sales Person"].isin(top6_sp)]
        fig4 = px.line(wk_sp6, x="Week", y="Inv", color="Sales Person",
                       color_discrete_sequence=C, markers=True)
        fig4.update_layout(**PL_general,
            xaxis=dict(tickangle=-40, tickfont=dict(size=9), title=""),
            yaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            legend=dict(orientation="h", x=0, y=1.15, font=dict(size=9)),
            height=320)
        st.plotly_chart(fig4, use_container_width=True)

    with r2:
        sh("🔄 Weekly Payment Status (Stacked)")
        fig5 = px.bar(wk_st_grp, x="Week", y="Count", color="Status",
                      barmode="stack",
                      color_discrete_map={
                          "Paid":         "#00c9a7",
                          "Accepted":     "#1a8fff",
                          "Not Accepted": "#ff6b35",
                      })
        fig5.update_layout(**PL_general,
            xaxis=dict(tickangle=-40, tickfont=dict(size=9), title=""),
            yaxis=dict(title="Submissions", gridcolor="#1a2a3a"),
            legend=dict(orientation="h", x=0, y=1.1, font=dict(size=9)),
            height=320)
        st.plotly_chart(fig5, use_container_width=True)

    st.markdown("---")

    # ── Row 5: Top Party weekly heatmap-style bar + line ─────────────────
    sh("👥 Weekly Invoice Value — Top 5 Parties (Line)")
    top5_p  = t_party.head(5)["Party Name"].tolist()
    wk_p5   = wk_party_top[wk_party_top["Party Name"].isin(top5_p)]
    fig6 = px.line(wk_p5, x="Week", y="Inv", color="Party Name",
                   color_discrete_sequence=C, markers=True)
    fig6.update_layout(**PL_general,
        xaxis=dict(tickangle=-40, tickfont=dict(size=9), title=""),
        yaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
        legend=dict(orientation="h", x=0, y=1.1, font=dict(size=9)),
        height=300)
    st.plotly_chart(fig6, use_container_width=True)

    st.markdown("---")

    # ── Row 6: Weekly detail table ────────────────────────────────────────
    sh("📋 Weekly Summary Table")
    wt = weekly[["Week","Count","Inv","LC","Paid_n","Paid_pct"]].copy()
    wt["Inv"]      = wt["Inv"].map(lambda x: f"${x:,.2f}")
    wt["LC"]       = wt["LC"].map(lambda x: f"${x:,.2f}")
    wt["Paid_pct"] = wt["Paid_pct"].map(lambda x: f"{x:.1f}%")
    wt["Paid_n"]   = wt["Paid_n"].astype(int)
    wt.columns     = ["Week","Submissions","Invoice Value (USD)","LC Value (USD)","Paid","Payment Rate"]
    st.dataframe(wt, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — FIRM & SALES PERSON
# ══════════════════════════════════════════════════════════════════════════════
with t3:
    l, r = st.columns(2)

    with l:
        sh("🏢 Firm-wise Invoice Value")
        fig = px.bar(by_firm, y="Firm Name", x="Inv", orientation="h",
                     color="Firm Name", color_discrete_sequence=C,
                     text=by_firm["Inv"].apply(usd))
        fig.update_traces(textposition="outside", textfont_size=10)
        fig.update_layout(**PL_general,
            xaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            yaxis=dict(title=""), showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

        sh("Firm Ranking Table")
        ft = by_firm.copy()
        ft["% Share"] = (ft["Inv"]/inv*100).map(lambda x: f"{x:.1f}%")
        ft["Inv"]     = ft["Inv"].map(lambda x: f"${x:,.2f}")
        ft.columns    = ["Firm","Invoice Value (USD)","Submissions","% Share"]
        st.dataframe(ft, use_container_width=True, hide_index=True)

    with r:
        sh("🧑‍💼 Sales Person — Invoice Value (Top 12)")
        fig2 = px.bar(spg.head(12), y="Sales Person", x="Inv", orientation="h",
                      color="Sales Person", color_discrete_sequence=C,
                      text=spg.head(12)["Inv"].apply(usd))
        fig2.update_traces(textposition="outside", textfont_size=10)
        fig2.update_layout(**PL_general,
            xaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            yaxis=dict(title=""), showlegend=False, height=400)
        st.plotly_chart(fig2, use_container_width=True)

        sh("Sales Person Ranking Table")
        st2 = spg[["Sales Person","Inv","N","Paid","Pct"]].copy()
        st2["Inv"]  = st2["Inv"].map(lambda x: f"${x:,.2f}")
        st2["Pct"]  = st2["Pct"].map(lambda x: f"{x:.1f}%")
        st2["Paid"] = st2["Paid"].astype(int)
        st2.columns = ["Sales Person","Invoice Value (USD)","Submissions","Paid","Payment Rate"]
        st.dataframe(st2, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — BANKS
# ══════════════════════════════════════════════════════════════════════════════
with t4:
    l, r = st.columns(2)
    with l:
        sh("🏦 Our Bank — Share (Donut)")
        fig = px.pie(by_bank, names="Our Bank", values="Inv",
                     color_discrete_sequence=C, hole=0.48)
        fig.update_layout(**PL)
        fig.update_traces(textinfo="label+percent", textfont_size=11)
        st.plotly_chart(fig, use_container_width=True)
    with r:
        sh("🏛️ Top 10 Party Banks")
        fig2 = px.bar(t_bname, y="Bank Name", x="Inv", orientation="h",
                      text=t_bname["Inv"].apply(usd), color_discrete_sequence=[C[4]])
        fig2.update_traces(textposition="outside", textfont_size=10, marker_color=C[4])
        fig2.update_layout(**PL, showlegend=False)

        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("---")
    a, b2 = st.columns(2)
    with a:
        sh("Our Bank Detail")
        ob = by_bank.copy()
        ob["% Share"] = (ob["Inv"]/inv*100).map(lambda x: f"{x:.1f}%")
        ob["Inv"]     = ob["Inv"].map(lambda x: f"${x:,.2f}")
        ob.columns    = ["Our Bank","Invoice Value (USD)","Submissions","% Share"]
        st.dataframe(ob, use_container_width=True, hide_index=True)
    with b2:
        sh("Top Party Banks Detail")
        pb = t_bname.copy()
        pb["% Share"] = (pb["Inv"]/inv*100).map(lambda x: f"{x:.1f}%")
        pb["Inv"]     = pb["Inv"].map(lambda x: f"${x:,.2f}")
        pb.columns    = ["Bank Name","Invoice Value (USD)","Submissions","% Share"]
        st.dataframe(pb, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — TOP PARTIES
# ══════════════════════════════════════════════════════════════════════════════
with t5:
    l, r = st.columns([3, 2])
    with l:
        sh("👥 Top 10 Party — Invoice Value")
        fig = px.bar(t_party, y="Party Name", x="Inv", orientation="h",
                     color="Party Name", color_discrete_sequence=C,
                     text=t_party["Inv"].apply(usd))
        fig.update_traces(textposition="outside", textfont_size=10)
        fig.update_layout(**PL_general,
            xaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            yaxis=dict(title="", autorange="reversed"),
            showlegend=False, height=400)
        st.plotly_chart(fig, use_container_width=True)
    with r:
        sh("Party Ranking")
        pt = t_party.copy()
        pt.insert(0, "#", range(1, len(pt)+1))
        pt["% Share"] = (pt["Inv"]/inv*100).map(lambda x: f"{x:.1f}%")
        pt["Inv"]     = pt["Inv"].map(lambda x: f"${x:,.2f}")
        pt            = pt[["#","Party Name","Inv","N","% Share"]]
        pt.columns    = ["#","Party","Invoice Value","Subs","% Share"]
        st.dataframe(pt, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 — PAYMENT STATUS
# ══════════════════════════════════════════════════════════════════════════════
with t6:
    s1, s2, s3 = st.columns(3)
    s1.metric("✅ Payment Received",       f"{paid_n:,}",  f"{usd(pay_v)} · {paid_n/N*100:.1f}%")
    s2.metric("⏳ Accepted — Pending Pmt", f"{acc_n:,}",   f"{usd(acc_v)} · {acc_n/N*100:.1f}%")
    s3.metric("❌ Not Accepted",           f"{nacc_n:,}",  f"{usd(nacc_v)} · {nacc_n/N*100:.1f}%")
    st.markdown("")

    st_df = pd.DataFrame({
        "Status": ["Payment Received","Accepted (Pending Pmt)","Not Accepted"],
        "Count":  [paid_n, acc_n, nacc_n],
        "Value":  [pay_v,  acc_v, nacc_v],
    })
    p1, p2 = st.columns(2)
    with p1:
        sh("By Count")
        fig = px.pie(st_df, names="Status", values="Count", hole=0.5,
                     color_discrete_sequence=["#00c9a7","#1a8fff","#ff6b35"])
        fig.update_layout(**PL)
        fig.update_traces(textinfo="label+percent", textfont_size=11)
        st.plotly_chart(fig, use_container_width=True)
    with p2:
        sh("By Invoice Value (USD)")
        fig2 = px.pie(st_df, names="Status", values="Value", hole=0.5,
                      color_discrete_sequence=["#00c9a7","#1a8fff","#ff6b35"])
        fig2.update_layout(**PL, showlegend=False)
        fig2.update_traces(textinfo="label+percent", textfont_size=11)
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("---")
    sh("Sales Person — Paid vs Not Yet Paid")
    fig3 = go.Figure()
    fig3.add_bar(name="Paid",         x=spg["Sales Person"], y=spg["Paid"],          marker_color="#00c9a7")
    fig3.add_bar(name="Not Yet Paid", x=spg["Sales Person"], y=spg["N"]-spg["Paid"], marker_color="#1a8fff")
    fig3.update_layout(**PL_general, barmode="stack",
        yaxis=dict(title="Submissions", gridcolor="#1a2a3a"),
        xaxis=dict(title="", tickangle=-35),
        legend=dict(orientation="h", x=0, y=1.1))
    st.plotly_chart(fig3, use_container_width=True)

    st.markdown("---")
    sh("Full Record Table with Search")
    search_text = st.text_input("🔍 Search all columns", key="global_search")
    pft = df.copy()
    pft["Status"] = pft.apply(lambda r:
        "✅ Paid" if pd.notna(r["Payment. Rcv Dt"]) else
        "⏳ Accepted" if pd.notna(r["Bank Accept Date"]) else
        "❌ Not Accepted", axis=1)
    if search_text:
        mask = pft.astype(str).apply(lambda c: c.str.contains(search_text, case=False, na=False)).any(axis=1)
        pft = pft[mask]

    date_columns = [
        "Bank Submition Date", "Bank Ref Date", "Lc Date",
        "Maturity Date", "Payment. Rcv Dt", "Date"
    ]
    for col in date_columns:
        if col in pft.columns:
            pft[col] = pd.to_datetime(pft[col], errors="coerce").dt.strftime("%d %b %Y")

    internal_columns = ["_date", "MonthSort", "Month", "WeekSort", "Week", "DayName"]
    report_column_order = [
        "Firm Name",
        "Our Bank",
        "Bank Submition Date",
        "Bank Ref Date",
        "Bank Refno",
        "Party Name",
        "LC No",
        "Lc Date",
        "Tenor",
        "Bank Name",
        "Invoice No",
        "Invoice Date",
        "Invoice Qty",
        "Invoice Value",
        "Bank Accept Date",
        "Maturity Date",
        "Payment. Rcv Dt",
        "Sales Person",
        "Week",
        "DayName",
        "Date",
        "Status",
    ]
    export_cols = [c for c in report_column_order if c in pft.columns]
    extra_cols = [c for c in pft.columns if c not in export_cols and c not in internal_columns]
    pft_export = pft[export_cols + extra_cols]
    pft_display = pft_export.copy()

    def make_column_widths(dataframe, min_width=100, max_width=450, char_width=8):
        text_df = dataframe.fillna("").astype(str)
        widths = {}
        max_lengths = text_df.apply(lambda col: col.str.len().max(), axis=0)
        for col, max_len in max_lengths.items():
            header_len = len(str(col))
            width = max(header_len, int(max_len or 0)) * char_width + 24
            widths[col] = min(max_width, max(min_width, width))
    
        # Ensure Bank Refno gets extra width to prevent forced wrapping in the display table
        if "Bank Refno" in widths:
            widths["Bank Refno"] = max(widths["Bank Refno"], 285)
        # Ensure Invoice Value column is wider so currency doesn't wrap as often
        if "Invoice Value" in widths:
            widths["Invoice Value"] = max(widths["Invoice Value"], 160)
        return widths

    # Compute total invoice value for the current (filtered & searched) dataset
    total_invoice_value = 0.0
    if "Invoice Value" in pft_export.columns:
        try:
            total_invoice_value = float(pft_export["Invoice Value"].sum())
        except Exception:
            total_invoice_value = 0.0

    # Format Invoice Value for display (USD)
    if "Invoice Value" in pft_display.columns:
        pft_display["Invoice Value"] = pft_display["Invoice Value"].map(lambda x: f"${x:,.2f}")
    column_widths = make_column_widths(pft_display)
    column_config = {
        col: st.column_config.TextColumn(width=column_widths[col])
        for col in pft_display.columns
    }
    st.dataframe(pft_display, use_container_width=True, hide_index=True,
                 height=500, column_config=column_config)
    # Show filtered total for Invoice Value
    if "Invoice Value" in pft_export.columns:
        st.markdown(f"**Total Invoice Value (filtered):** ${total_invoice_value:,.2f}")

    def table_widths_to_pdf_inches(widths, min_width=1.2, max_width=4.5):
        return [max(min_width, min(max_width, w / 96.0)) for w in widths]

    table_pdf_widths = table_widths_to_pdf_inches([column_widths[col] for col in pft_display.columns])

    

    col_csv, col_pdf = st.columns([1, 1])
    col_csv.download_button("📥 Download CSV",
        pft_export.to_csv(index=False).encode("utf-8"),
        "bank_submit_filtered.csv", "text/csv")

    # PDF width controls: let user choose auto, equal or custom per-column widths
    pdf_width_mode = st.selectbox("PDF column width mode", [
        "Auto (by content)", "Equal", "Custom (inches, comma-separated)"
    ])
    custom_widths_input = ""
    if pdf_width_mode == "Custom (inches, comma-separated)":
        custom_widths_input = st.text_input("Custom widths (comma-separated, e.g. 1.0,2.5,1.5)")

    if isinstance(date_range, tuple) and len(date_range) == 2:
        pdf_period = f"{date_range[0].strftime('%d %b %Y')} – {date_range[1].strftime('%d %b %Y')}"
    else:
        pdf_period = str(date_range)

    if REPORTLAB_AVAILABLE:
        def df_to_pdf_bytes(df, title="Bank submit status", subtitle="", custom_widths=None, generated_at=""):
            buf = BytesIO()
            left_margin = right_margin = top_margin = bottom_margin = 24
            page_width, page_height = landscape(A3)
            usable_width = page_width - left_margin - right_margin
            df_text = df.fillna("").astype(str)
            header_font = "Helvetica-Bold"
            cell_font = "Helvetica"
            header_font_size = 10
            cell_font_size = 8
            min_col_width = 1.2 * inch
            max_col_width = 4.5 * inch

            def measure_text_width(text: str, font: str, size: float) -> float:
                return pdfmetrics.stringWidth(str(text), font, size)

            # If caller supplied custom_widths (list of inches) or asked for equal widths,
            # build col_widths from that; otherwise fall back to auto-measurement.
            col_widths = []
            if custom_widths and isinstance(custom_widths, (list, tuple)):
                # convert inches to points
                col_widths = [max(min_col_width, min(max_col_width, float(w) * inch)) for w in custom_widths]
                # if fewer widths than columns, distribute remaining space equally
                if len(col_widths) < len(df_text.columns):
                    remaining = len(df_text.columns) - len(col_widths)
                    remaining_space = max(0, usable_width - sum(col_widths))
                    add_w = remaining_space / remaining if remaining else min_col_width
                    col_widths.extend([max(min_col_width, min(max_col_width, add_w)) for _ in range(remaining)])
                total_width = sum(col_widths)
                if total_width > usable_width and total_width > 0:
                    scale = usable_width / total_width
                    col_widths = [w * scale for w in col_widths]
            elif custom_widths == "EQUAL":
                eq = usable_width / len(df_text.columns) if len(df_text.columns) else usable_width
                col_widths = [eq for _ in df_text.columns]
            else:
                sample_count = 250
                for col in df_text.columns:
                    header_width = measure_text_width(col, header_font, header_font_size)
                    values = df_text[col].tolist()
                    if len(values) > sample_count:
                        step = max(1, len(values) // sample_count)
                        values = values[::step]
                    measured = [measure_text_width(v, cell_font, cell_font_size) for v in values if v is not None]
                    if measured:
                        sorted_widths = sorted(measured)
                        percentile_index = min(len(sorted_widths) - 1, int(len(sorted_widths) * 0.9))
                        cell_width = sorted_widths[percentile_index]
                    else:
                        cell_width = measure_text_width("M", cell_font, cell_font_size)
                    width = max(header_width, cell_width) + 16
                    col_widths.append(min(max_col_width, max(min_col_width, width)))

                total_width = sum(col_widths)
                if total_width > usable_width and total_width > 0:
                    scale = usable_width / total_width
                    col_widths = [w * scale for w in col_widths]
                elif total_width < usable_width and total_width > 0:
                    extra = usable_width - total_width
                    col_widths = [w + extra * (w / total_width) for w in col_widths]

            styles = getSampleStyleSheet()
            header_style = ParagraphStyle(
                name="HeaderStyle",
                parent=styles["Normal"],
                fontName="Helvetica-Bold",
                fontSize=10,
                leading=11,
                textColor=colors.white,
                alignment=TA_LEFT,
                spaceAfter=2,
                spaceBefore=2,
                wordWrap="CJK",
            )
            cell_style = ParagraphStyle(
                name="CellStyle",
                parent=styles["Normal"],
                fontName="Helvetica",
                fontSize=8,
                leading=10,
                alignment=TA_LEFT,
                wordWrap="CJK",
                allowWidows=True,
                allowOrphans=True,
                spaceBefore=0,
                spaceAfter=0,
                leftIndent=0,
                rightIndent=0,
            )

            def chunk_columns(columns, widths, max_width):
                groups = []
                current_cols = []
                current_width = 0.0
                for col, w in zip(columns, widths):
                    if current_cols and current_width + w > max_width:
                        groups.append(current_cols)
                        current_cols = [col]
                        current_width = w
                    else:
                        current_cols.append(col)
                        current_width += w
                if current_cols:
                    groups.append(current_cols)
                return groups

            data_pages = []
            column_groups = chunk_columns(list(df.columns), col_widths, usable_width)
            style = TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1a8fff")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("WORDWRAP", (0, 0), (-1, -1), "CJK"),
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey]),
            ])

            for group_index, group_cols in enumerate(column_groups):
                group_widths = [col_widths[list(df.columns).index(col)] for col in group_cols]
                page_data = [[Paragraph(str(col), header_style) for col in group_cols]]
                for row in df.fillna("").astype(str).values.tolist():
                    page_data.append([Paragraph(str(row[list(df.columns).index(col)]), cell_style) for col in group_cols])

                page_table = LongTable(
                    page_data,
                    repeatRows=1,
                    colWidths=group_widths,
                    hAlign="LEFT",
                    splitByRow=1,
                    spaceBefore=12,
                    spaceAfter=12,
                )
                page_table.setStyle(style)
                data_pages.append(page_table)
                if group_index < len(column_groups) - 1:
                    data_pages.append(PageBreak())

            def page_header(canvas, doc):
                canvas.saveState()
                title_y = page_height - top_margin + 10
                subtitle_y = title_y - 16
                generated_y = subtitle_y - 14
                canvas.setFont("Helvetica-Bold", 20)
                canvas.drawCentredString(page_width / 2, title_y, str(title))
                canvas.setFont("Helvetica", 11)
                canvas.drawCentredString(page_width / 2, subtitle_y, str(subtitle))
                if generated_at:
                    canvas.setFont("Helvetica", 8)
                    canvas.drawString(left_margin, generated_y, f"Generated on: {generated_at}")
                canvas.setFont("Helvetica", 8)
                canvas.drawRightString(page_width - right_margin, title_y, f"Page {doc.page}")
                # Footer branding text
                canvas.setFont("Helvetica-Bold", 8)
                canvas.drawRightString(page_width - right_margin, bottom_margin / 2, "ASM")
                canvas.restoreState()

            doc = SimpleDocTemplate(
                buf,
                pagesize=landscape(A3),
                leftMargin=left_margin,
                rightMargin=right_margin,
                topMargin=top_margin + 28,
                bottomMargin=bottom_margin,
            )
            doc.build(data_pages, onFirstPage=page_header, onLaterPages=page_header)
            buf.seek(0)
            return buf.read()

        try:
            # determine parsed_custom based on UI selection
            parsed_custom = None
            if pdf_width_mode == "Auto (by content)":
                parsed_custom = table_pdf_widths
            elif pdf_width_mode == "Equal":
                parsed_custom = "EQUAL"
            elif pdf_width_mode == "Custom (inches, comma-separated)" and custom_widths_input:
                try:
                    parsed_custom = [float(x.strip()) for x in custom_widths_input.split(",") if x.strip()]
                except Exception:
                    col_pdf.error("Invalid custom widths. Use comma-separated numbers like: 1.0,2.5,1.5")
                    parsed_custom = None

            generated_at = datetime.now().strftime("%d %b %Y %H:%M:%S")
            # Prepare PDF dataframe: format Invoice Value as USD and append a totals row
            pdf_df = pft_export.copy()
            if "Invoice Value" in pdf_df.columns:
                pdf_df["Invoice Value"] = pdf_df["Invoice Value"].map(lambda x: f"${x:,.2f}")
            # Append a totals row (label in first column, total in Invoice Value column)
            try:
                total_val = float(total_invoice_value)
            except Exception:
                total_val = 0.0
            if len(pdf_df.columns) > 0:
                totals_row = {c: "" for c in pdf_df.columns}
                totals_row[pdf_df.columns[0]] = "TOTAL"
                if "Invoice Value" in pdf_df.columns:
                    totals_row["Invoice Value"] = f"${total_val:,.2f}"
                pdf_df = pd.concat([pdf_df, pd.DataFrame([totals_row])], ignore_index=True)

            pdf_bytes = df_to_pdf_bytes(
                pdf_df,
                subtitle=pdf_period,
                custom_widths=parsed_custom,
                generated_at=generated_at,
            )
            col_pdf.download_button("📄 Download PDF",
                pdf_bytes,
                "bank_submit_status.pdf",
                "application/pdf")
        except Exception as e:
            col_pdf.error(f"Could not generate PDF: {e}")
    else:
        col_pdf.warning("Install reportlab to enable PDF export: pip install reportlab")


# ─────────────────────────────────────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    f"<p style='text-align:center;color:#1a2a3a;font-size:11px;letter-spacing:.1em'>"
    f" Asm@2026 BANK SUBMIT HISTORY DASHBOARD &nbsp;·&nbsp; {N:,} RECORDS &nbsp;·&nbsp; {period}"
    f"</p>", unsafe_allow_html=True)
