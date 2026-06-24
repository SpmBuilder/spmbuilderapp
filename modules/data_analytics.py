"""
modules/data_analytics.py — Descriptive stats, distributions, correlations, and charts
"""

import streamlit as st
import pandas as pd
import numpy as np
from config.settings import page_header

# ── Helpers ───────────────────────────────────────────────────────────────────

def _get_datasets() -> dict:
    return st.session_state.get("uploaded_datasets", {})

def _dataset_names() -> list[str]:
    names = list(_get_datasets().keys())
    if "db_query_result" in st.session_state:
        names = ["[DB Query Result]"] + names
    return names

def _resolve(name: str) -> pd.DataFrame | None:
    if name == "[DB Query Result]":
        return st.session_state.get("db_query_result")
    return _get_datasets().get(name, {}).get("df")

def _numeric_cols(df: pd.DataFrame) -> list[str]:
    return df.select_dtypes(include="number").columns.tolist()

def _cat_cols(df: pd.DataFrame) -> list[str]:
    return df.select_dtypes(include=["object", "category", "bool"]).columns.tolist()


# ── Tab: Overview ─────────────────────────────────────────────────────────────

def _tab_overview():
    names = _dataset_names()
    if not names:
        st.info("No datasets available. Upload data first.")
        return

    name = st.selectbox("Dataset", names, key="ana_ds")
    df = _resolve(name)
    if df is None:
        return

    num_cols = _numeric_cols(df)
    cat_cols = _cat_cols(df)

    # Top metrics
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Rows",        f"{len(df):,}")
    c2.metric("Columns",     len(df.columns))
    c3.metric("Numeric",     len(num_cols))
    c4.metric("Categorical", len(cat_cols))

    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

    # Describe
    if num_cols:
        st.markdown("**Numeric summary**")
        desc = df[num_cols].describe().T.reset_index().rename(columns={"index": "column"})
        desc = desc.round(4)
        st.dataframe(desc, use_container_width=True)

    if cat_cols:
        st.markdown("**Categorical summary**")
        rows = []
        for c in cat_cols:
            rows.append({
                "Column":  c,
                "Unique":  df[c].nunique(),
                "Top":     str(df[c].mode().iloc[0]) if not df[c].mode().empty else "—",
                "Top Freq": int(df[c].value_counts().iloc[0]) if not df[c].empty else 0,
                "Null %":  f"{df[c].isna().mean()*100:.1f}%",
            })
        st.dataframe(pd.DataFrame(rows), use_container_width=True)

    # Null heatmap (text-based)
    st.markdown("**Null density by column**")
    null_pct = (df.isna().mean() * 100).reset_index()
    null_pct.columns = ["Column", "Null %"]
    null_pct = null_pct[null_pct["Null %"] > 0].sort_values("Null %", ascending=False)
    if null_pct.empty:
        st.success("No nulls found across any column.")
    else:
        st.dataframe(null_pct, use_container_width=True)


# ── Tab: Charts ───────────────────────────────────────────────────────────────

def _tab_charts():
    names = _dataset_names()
    if not names:
        st.info("No datasets available.")
        return

    name = st.selectbox("Dataset", names, key="chart_ds")
    df = _resolve(name)
    if df is None:
        return

    num_cols = _numeric_cols(df)
    cat_cols = _cat_cols(df)

    chart_type = st.selectbox(
        "Chart type",
        ["Histogram", "Bar / Value counts", "Line", "Scatter", "Box"],
        key="chart_type",
    )

    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

    if chart_type == "Histogram":
        if not num_cols:
            st.warning("No numeric columns in this dataset.")
            return
        col   = st.selectbox("Column", num_cols, key="hist_col")
        bins  = st.slider("Bins", 5, 100, 30, key="hist_bins")
        series = df[col].dropna()
        counts, edges = np.histogram(series, bins=bins)
        hist_df = pd.DataFrame({"bin": edges[:-1].round(4), "count": counts})
        st.bar_chart(hist_df.set_index("bin"), use_container_width=True)

    elif chart_type == "Bar / Value counts":
        col = st.selectbox(
            "Column",
            cat_cols + num_cols,
            key="bar_col",
        )
        top_n = st.slider("Top N values", 5, 50, 15, key="bar_topn")
        vc = df[col].astype(str).value_counts().head(top_n)
        st.bar_chart(vc, use_container_width=True)

    elif chart_type == "Line":
        if not num_cols:
            st.warning("No numeric columns.")
            return
        x_col = st.selectbox("X-axis", df.columns.tolist(), key="line_x")
        y_cols = st.multiselect("Y-axis (one or more numeric)", num_cols, default=num_cols[:1], key="line_y")
        if y_cols:
            plot_df = df[[x_col] + y_cols].set_index(x_col).dropna()
            st.line_chart(plot_df, use_container_width=True)

    elif chart_type == "Scatter":
        if len(num_cols) < 2:
            st.warning("Need at least 2 numeric columns.")
            return
        x = st.selectbox("X",      num_cols,          key="sc_x")
        y = st.selectbox("Y",      num_cols, index=1, key="sc_y")
        plot_df = df[[x, y]].dropna().rename(columns={x: "x", y: "y"})
        st.scatter_chart(plot_df, x="x", y="y", use_container_width=True)

    elif chart_type == "Box":
        if not num_cols:
            st.warning("No numeric columns.")
            return
        cols = st.multiselect("Columns", num_cols, default=num_cols[:3], key="box_cols")
        if cols:
            desc = df[cols].describe().loc[["min", "25%", "50%", "75%", "max"]].T
            st.dataframe(desc.round(4), use_container_width=True)
            st.caption("Box plots rendered as quartile table (native Streamlit charts don't support box plots).")


# ── Tab: Correlation ──────────────────────────────────────────────────────────

def _tab_correlation():
    names = _dataset_names()
    if not names:
        st.info("No datasets available.")
        return

    name = st.selectbox("Dataset", names, key="corr_ds")
    df = _resolve(name)
    if df is None:
        return

    num_cols = _numeric_cols(df)
    if len(num_cols) < 2:
        st.warning("Need at least 2 numeric columns for correlation.")
        return

    method = st.radio("Method", ["pearson", "spearman", "kendall"], horizontal=True, key="corr_method")
    cols   = st.multiselect("Columns", num_cols, default=num_cols[:8], key="corr_cols")

    if len(cols) < 2:
        st.info("Select at least 2 columns.")
        return

    corr = df[cols].corr(method=method).round(4)
    st.dataframe(corr.style.background_gradient(cmap="Greys"), use_container_width=True)

    # Top pairs
    st.markdown("**Top correlated pairs**")
    pairs = []
    for i in range(len(cols)):
        for j in range(i + 1, len(cols)):
            pairs.append({
                "Column A":    cols[i],
                "Column B":    cols[j],
                "Correlation": round(corr.iloc[i, j], 4),
            })
    pairs_df = pd.DataFrame(pairs).sort_values("Correlation", key=abs, ascending=False)
    st.dataframe(pairs_df, use_container_width=True)


# ── Tab: Pivot ────────────────────────────────────────────────────────────────

def _tab_pivot():
    names = _dataset_names()
    if not names:
        st.info("No datasets available.")
        return

    name = st.selectbox("Dataset", names, key="pivot_ds")
    df = _resolve(name)
    if df is None:
        return

    num_cols = _numeric_cols(df)
    all_cols = df.columns.tolist()

    c1, c2, c3 = st.columns(3)
    index_col  = c1.selectbox("Rows (index)",   all_cols,          key="pv_index")
    col_col    = c2.selectbox("Columns",         all_cols, index=min(1, len(all_cols)-1), key="pv_col")
    value_col  = c3.selectbox("Values",          num_cols if num_cols else all_cols, key="pv_val")
    agg_fn     = st.selectbox("Aggregation", ["sum", "mean", "count", "min", "max"], key="pv_agg")

    if st.button("Build Pivot", type="primary"):
        try:
            pivot = df.pivot_table(
                index=index_col, columns=col_col, values=value_col,
                aggfunc=agg_fn, fill_value=0,
            )
            st.dataframe(pivot.round(4), use_container_width=True)

            # Store for report generation
            st.session_state["last_pivot"] = pivot

            csv = pivot.to_csv().encode()
            st.download_button("⬇ Download Pivot CSV", csv, "pivot.csv", "text/csv")
        except Exception as e:
            st.error(str(e))


# ── Main render ───────────────────────────────────────────────────────────────

def render():
    page_header(
        "Analytics",
        "Explore distributions, correlations, charts, and pivot tables.",
    )
    tabs = st.tabs(["📋  Overview", "📊  Charts", "🔗  Correlation", "📐  Pivot"])
    with tabs[0]: _tab_overview()
    with tabs[1]: _tab_charts()
    with tabs[2]: _tab_correlation()
    with tabs[3]: _tab_pivot()
