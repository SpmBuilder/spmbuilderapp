"""
modules/data_uploader.py — Upload data from CSV, Excel, JSON, Parquet, and URLs
"""

import io
import json
import streamlit as st
import pandas as pd
from config.settings import page_header, badge

# ── Helpers ───────────────────────────────────────────────────────────────────

SUPPORTED_EXTENSIONS = {
    "csv":     "text/csv",
    "xlsx":    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "xls":     "application/vnd.ms-excel",
    "json":    "application/json",
    "parquet": "application/octet-stream",
    "tsv":     "text/tab-separated-values",
}

def _load_file(file) -> pd.DataFrame | None:
    ext = file.name.rsplit(".", 1)[-1].lower()
    try:
        if ext == "csv":
            return pd.read_csv(file)
        elif ext == "tsv":
            return pd.read_csv(file, sep="\t")
        elif ext in ("xlsx", "xls"):
            return pd.read_excel(file)
        elif ext == "json":
            data = json.load(file)
            return pd.json_normalize(data) if isinstance(data, (list, dict)) else None
        elif ext == "parquet":
            return pd.read_parquet(file)
    except Exception as e:
        st.error(f"Could not parse `{file.name}`: {e}")
    return None


def _load_url(url: str) -> pd.DataFrame | None:
    try:
        import urllib.request
        with urllib.request.urlopen(url, timeout=15) as resp:
            content = resp.read()
        ext = url.rsplit(".", 1)[-1].lower().split("?")[0]
        if ext in ("csv", "tsv"):
            sep = "\t" if ext == "tsv" else ","
            return pd.read_csv(io.BytesIO(content), sep=sep)
        elif ext in ("xlsx", "xls"):
            return pd.read_excel(io.BytesIO(content))
        elif ext == "json":
            data = json.loads(content)
            return pd.json_normalize(data) if isinstance(data, (list, dict)) else None
        else:
            # Try CSV as a fallback
            return pd.read_csv(io.BytesIO(content))
    except Exception as e:
        st.error(f"Could not fetch URL: {e}")
    return None


def _profile(df: pd.DataFrame) -> pd.DataFrame:
    """Quick column profile."""
    rows = []
    for col in df.columns:
        s = df[col]
        rows.append({
            "Column":    col,
            "Type":      str(s.dtype),
            "Non-Null":  int(s.notna().sum()),
            "Null":      int(s.isna().sum()),
            "Null %":    f"{s.isna().mean()*100:.1f}%",
            "Unique":    int(s.nunique()),
            "Sample":    str(s.dropna().iloc[0]) if s.notna().any() else "—",
        })
    return pd.DataFrame(rows)


def _store_dataset(name: str, df: pd.DataFrame, source: str):
    store = st.session_state.setdefault("uploaded_datasets", {})
    store[name] = {"df": df, "source": source, "rows": len(df), "cols": len(df.columns)}


# ── Tab: File upload ──────────────────────────────────────────────────────────

def _tab_file():
    st.markdown("##### Upload one or more files")
    files = st.file_uploader(
        "Drag and drop or click to browse",
        type=list(SUPPORTED_EXTENSIONS.keys()),
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    if not files:
        st.markdown(
            """
            <div style="text-align:center;padding:2rem;color:#444;font-size:0.82rem;">
                Supported: CSV · TSV · XLSX · XLS · JSON · Parquet
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    for file in files:
        with st.expander(f"📄  {file.name}  —  {file.size / 1024:.1f} KB", expanded=True):
            df = _load_file(file)
            if df is None:
                continue

            c1, c2, c3 = st.columns(3)
            c1.metric("Rows", f"{len(df):,}")
            c2.metric("Columns", len(df.columns))
            c3.metric("Memory", f"{df.memory_usage(deep=True).sum() / 1024:.1f} KB")

            inner = st.tabs(["Preview", "Profile", "Raw"])

            with inner[0]:
                st.dataframe(df.head(50), use_container_width=True)

            with inner[1]:
                st.dataframe(_profile(df), use_container_width=True)

            with inner[2]:
                st.code(df.head(5).to_string(), language="text")

            alias = st.text_input(
                "Save as dataset name",
                value=file.name.rsplit(".", 1)[0],
                key=f"alias_{file.name}",
            )
            if st.button("Save to workspace", key=f"save_{file.name}", type="primary"):
                _store_dataset(alias, df, f"file:{file.name}")
                st.success(f"Saved as **{alias}**")


# ── Tab: URL fetch ────────────────────────────────────────────────────────────

def _tab_url():
    st.markdown("##### Fetch data from a URL")
    url = st.text_input("URL", placeholder="https://example.com/data.csv")
    alias = st.text_input("Dataset name", placeholder="my_remote_data")

    if st.button("Fetch", type="primary", disabled=not url.strip()):
        if not alias.strip():
            st.warning("Please provide a dataset name.")
            return
        with st.spinner("Fetching…"):
            df = _load_url(url.strip())
        if df is not None:
            _store_dataset(alias.strip(), df, f"url:{url}")
            st.success(f"Fetched {len(df):,} rows × {len(df.columns)} columns → saved as **{alias}**")
            st.dataframe(df.head(20), use_container_width=True)


# ── Tab: Paste / manual ───────────────────────────────────────────────────────

def _tab_paste():
    st.markdown("##### Paste CSV or JSON text")
    fmt = st.radio("Format", ["CSV", "JSON"], horizontal=True)
    raw = st.text_area("Paste data here", height=200, placeholder="col1,col2,col3\n1,2,3\n4,5,6")
    alias = st.text_input("Dataset name", placeholder="pasted_data")

    if st.button("Parse & Save", type="primary", disabled=not raw.strip()):
        if not alias.strip():
            st.warning("Please provide a dataset name.")
            return
        try:
            if fmt == "CSV":
                df = pd.read_csv(io.StringIO(raw))
            else:
                data = json.loads(raw)
                df = pd.json_normalize(data)
            _store_dataset(alias.strip(), df, "paste")
            st.success(f"Parsed {len(df):,} rows × {len(df.columns)} columns → **{alias}**")
            st.dataframe(df, use_container_width=True)
        except Exception as e:
            st.error(f"Parse error: {e}")


# ── Tab: Workspace ────────────────────────────────────────────────────────────

def _tab_workspace():
    store = st.session_state.get("uploaded_datasets", {})
    if not store:
        st.info("No datasets in the workspace yet.")
        return

    for name, meta in list(store.items()):
        with st.expander(f"📦  {name}  ·  {meta['rows']:,} rows × {meta['cols']} cols", expanded=False):
            col1, col2 = st.columns([5, 1])
            col1.caption(f"Source: `{meta['source']}`")
            if col2.button("Remove", key=f"rm_{name}"):
                del store[name]
                st.rerun()

            df = meta["df"]
            st.dataframe(df.head(30), use_container_width=True)

            csv_bytes = df.to_csv(index=False).encode()
            st.download_button(
                "⬇ Download CSV",
                csv_bytes,
                f"{name}.csv",
                "text/csv",
                key=f"dl_{name}",
            )


# ── Main render ───────────────────────────────────────────────────────────────

def render():
    page_header(
        "Data Upload",
        "Import data from files, remote URLs, or paste directly.",
    )

    tabs = st.tabs(["📁  File Upload", "🌐  URL Fetch", "📋  Paste", "📦  Workspace"])
    with tabs[0]: _tab_file()
    with tabs[1]: _tab_url()
    with tabs[2]: _tab_paste()
    with tabs[3]: _tab_workspace()
