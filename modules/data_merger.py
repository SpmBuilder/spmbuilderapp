"""
modules/data_merger.py — Join, union, and blend datasets from the workspace
"""

import streamlit as st
import pandas as pd
from config.settings import page_header

# ── Helpers ───────────────────────────────────────────────────────────────────

def _get_datasets() -> dict:
    return st.session_state.get("uploaded_datasets", {})


def _dataset_names() -> list[str]:
    names = list(_get_datasets().keys())
    # Also expose DB query result if present
    if "db_query_result" in st.session_state:
        names = ["[DB Query Result]"] + names
    return names


def _resolve(name: str) -> pd.DataFrame | None:
    if name == "[DB Query Result]":
        return st.session_state.get("db_query_result")
    return _get_datasets().get(name, {}).get("df")


def _store_result(name: str, df: pd.DataFrame, source: str):
    store = st.session_state.setdefault("uploaded_datasets", {})
    store[name] = {"df": df, "source": source, "rows": len(df), "cols": len(df.columns)}


# ── Tab: Join ─────────────────────────────────────────────────────────────────

def _tab_join():
    st.markdown("##### Join two datasets (SQL-style)")
    names = _dataset_names()
    if len(names) < 2:
        st.info("You need at least 2 datasets in the workspace. Use the Data Upload module to add them.")
        return

    c1, c2 = st.columns(2)
    left_name  = c1.selectbox("Left dataset",  names, key="join_left")
    right_name = c2.selectbox("Right dataset", names, index=1, key="join_right")

    df_l = _resolve(left_name)
    df_r = _resolve(right_name)

    if df_l is None or df_r is None:
        st.error("Could not load one of the datasets.")
        return

    common_cols = sorted(set(df_l.columns) & set(df_r.columns))

    c3, c4 = st.columns(2)
    left_key  = c3.selectbox("Left key column",  df_l.columns.tolist(), key="join_lkey")
    right_key = c4.selectbox("Right key column", df_r.columns.tolist(),
                             index=df_r.columns.tolist().index(left_key)
                             if left_key in df_r.columns else 0,
                             key="join_rkey")

    join_type = st.radio(
        "Join type",
        ["inner", "left", "right", "outer"],
        horizontal=True,
        key="join_type",
    )

    suffixes_l = st.text_input("Left suffix (for duplicate columns)", value="_left",  key="suf_l")
    suffixes_r = st.text_input("Right suffix",                        value="_right", key="suf_r")

    out_name = st.text_input("Save result as", value=f"{left_name}_join_{right_name}", key="join_out")

    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

    col_preview = st.columns(2)
    with col_preview[0]:
        st.caption(f"**{left_name}** — {len(df_l):,} rows, {len(df_l.columns)} cols")
        st.dataframe(df_l.head(5), use_container_width=True)
    with col_preview[1]:
        st.caption(f"**{right_name}** — {len(df_r):,} rows, {len(df_r.columns)} cols")
        st.dataframe(df_r.head(5), use_container_width=True)

    if st.button("Run Join", type="primary"):
        if not out_name.strip():
            st.warning("Please set an output dataset name.")
            return
        with st.spinner("Merging…"):
            try:
                result = pd.merge(
                    df_l, df_r,
                    left_on=left_key, right_on=right_key,
                    how=join_type,
                    suffixes=(suffixes_l, suffixes_r),
                )
                _store_result(out_name.strip(), result, f"join:{left_name}+{right_name}")
                st.success(f"Merged → **{out_name}** — {len(result):,} rows × {len(result.columns)} cols")
                st.dataframe(result.head(50), use_container_width=True)
            except Exception as e:
                st.error(str(e))


# ── Tab: Union ────────────────────────────────────────────────────────────────

def _tab_union():
    st.markdown("##### Stack datasets vertically (UNION)")
    names = _dataset_names()
    if len(names) < 2:
        st.info("Need at least 2 datasets in the workspace.")
        return

    selected = st.multiselect("Datasets to union (in order)", names, key="union_sel")
    ignore_index = st.checkbox("Reset row index", value=True, key="union_reset")
    match_cols = st.checkbox("Only include common columns", value=False, key="union_common")
    out_name = st.text_input("Save result as", value="union_result", key="union_out")

    if len(selected) < 2:
        st.info("Select at least 2 datasets.")
        return

    frames = []
    for n in selected:
        df = _resolve(n)
        if df is None:
            st.warning(f"Could not resolve **{n}**.")
            return
        frames.append(df)

    if match_cols:
        common = set(frames[0].columns)
        for f in frames[1:]:
            common &= set(f.columns)
        frames = [f[list(common)] for f in frames]

    if st.button("Run Union", type="primary"):
        if not out_name.strip():
            st.warning("Please set an output dataset name.")
            return
        with st.spinner("Stacking…"):
            try:
                result = pd.concat(frames, ignore_index=ignore_index)
                _store_result(out_name.strip(), result, f"union:{'+'.join(selected)}")
                st.success(f"Union → **{out_name}** — {len(result):,} rows × {len(result.columns)} cols")
                st.dataframe(result.head(30), use_container_width=True)
            except Exception as e:
                st.error(str(e))


# ── Tab: Column mapping / rename ──────────────────────────────────────────────

def _tab_reshape():
    st.markdown("##### Reshape: select, rename, and reorder columns")
    names = _dataset_names()
    if not names:
        st.info("No datasets in the workspace.")
        return

    src_name = st.selectbox("Source dataset", names, key="reshape_src")
    df = _resolve(src_name)
    if df is None:
        return

    st.caption(f"{len(df):,} rows × {len(df.columns)} columns")

    # Build a rename map UI
    st.markdown("**Column selection & rename**")
    col_cfg = {}
    for i, col in enumerate(df.columns):
        c1, c2, c3 = st.columns([0.5, 3, 3])
        keep    = c1.checkbox("", value=True, key=f"rsh_keep_{i}")
        orig    = c2.text_input("Original", value=col, disabled=True, key=f"rsh_orig_{i}")
        renamed = c3.text_input("Rename to", value=col, key=f"rsh_new_{i}")
        if keep:
            col_cfg[col] = renamed

    out_name = st.text_input("Save result as", value=f"{src_name}_reshaped", key="reshape_out")

    if st.button("Apply", type="primary"):
        if not col_cfg:
            st.warning("Select at least one column.")
            return
        result = df[list(col_cfg.keys())].rename(columns=col_cfg)
        _store_result(out_name.strip() or "reshaped", result, f"reshape:{src_name}")
        st.success(f"Saved **{out_name}** — {len(result.columns)} columns selected.")
        st.dataframe(result.head(30), use_container_width=True)


# ── Main render ───────────────────────────────────────────────────────────────

def render():
    page_header(
        "Data Merger",
        "Join, union, and reshape datasets from your workspace.",
    )
    tabs = st.tabs(["🔗  Join", "📚  Union / Stack", "✏️  Reshape"])
    with tabs[0]: _tab_join()
    with tabs[1]: _tab_union()
    with tabs[2]: _tab_reshape()
