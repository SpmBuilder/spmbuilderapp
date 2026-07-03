"""
modules/report_generator.py — Generate PDF/Excel/CSV reports from datasets and templates
"""

import io
import json
import datetime
import streamlit as st
import pandas as pd
from config.settings import page_header

# ── Helpers ───────────────────────────────────────────────────────────────────

def _get_datasets() -> dict:
    return st.session_state.get("uploaded_datasets", {})

def _dataset_names() -> list[str]:
    names = list(_get_datasets().keys())
    if "db_query_result" in st.session_state:
        names = ["[DB Query Result]"] + names
    if "last_mapped" in st.session_state:
        names = ["[Last Mapped Output]"] + names
    return names

def _resolve(name: str) -> pd.DataFrame | None:
    if name == "[DB Query Result]":
        return st.session_state.get("db_query_result")
    if name == "[Last Mapped Output]":
        return st.session_state.get("last_mapped")
    return _get_datasets().get(name, {}).get("df")


# ── Report section builder ────────────────────────────────────────────────────

def _build_summary_section(df: pd.DataFrame) -> dict:
    num_cols = df.select_dtypes(include="number").columns.tolist()
    summary = {
        "row_count":  len(df),
        "col_count":  len(df.columns),
        "null_total": int(df.isna().sum().sum()),
    }
    if num_cols:
        desc = df[num_cols].describe().T
        summary["numeric_summary"] = desc.round(4).to_dict()
    return summary


# ── Tab: CSV / Excel export ───────────────────────────────────────────────────

def _tab_export():
    st.markdown("##### Export dataset to file")
    names = _dataset_names()
    if not names:
        st.info("No datasets in the workspace. Upload or query data first.")
        return

    name = st.selectbox("Dataset", names, key="rpt_exp_ds")
    df = _resolve(name)
    if df is None:
        return

    c1, c2, c3 = st.columns(3)
    c1.metric("Rows",    f"{len(df):,}")
    c2.metric("Columns", len(df.columns))
    c3.metric("Nulls",   f"{df.isna().sum().sum():,}")

    fmt = st.radio("Export format", ["CSV", "Excel (.xlsx)", "JSON", "TSV"], horizontal=True, key="rpt_fmt")

    # Column selection
    sel_cols = st.multiselect(
        "Columns to export (all if empty)",
        df.columns.tolist(),
        key="rpt_cols",
    )
    export_df = df[sel_cols] if sel_cols else df

    # Filter by row count
    max_rows = st.number_input(
        "Row limit (0 = all)",
        min_value=0, max_value=len(df), value=0, step=100,
        key="rpt_rowlimit",
    )
    if max_rows > 0:
        export_df = export_df.head(int(max_rows))

    filename_base = st.text_input("File name (without extension)", value=name.replace(" ", "_"), key="rpt_fname")

    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

    if st.button("Generate & Download", type="primary"):
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        fname_base = (filename_base.strip() or "export") + f"_{ts}"

        if fmt == "CSV":
            data  = export_df.to_csv(index=False).encode("utf-8")
            fname = f"{fname_base}.csv"
            mime  = "text/csv"

        elif fmt == "TSV":
            data  = export_df.to_csv(index=False, sep="\t").encode("utf-8")
            fname = f"{fname_base}.tsv"
            mime  = "text/tab-separated-values"

        elif fmt == "JSON":
            data  = export_df.to_json(orient="records", indent=2).encode("utf-8")
            fname = f"{fname_base}.json"
            mime  = "application/json"

        elif fmt == "Excel (.xlsx)":
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                export_df.to_excel(writer, index=False, sheet_name="Data")
            data  = buf.getvalue()
            fname = f"{fname_base}.xlsx"
            mime  = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        st.download_button(
            f"⬇  Download {fname}",
            data=data,
            file_name=fname,
            mime=mime,
            key="rpt_dl_btn",
        )
        st.success(f"Ready to download: **{fname}** — {len(export_df):,} rows")


# ── Tab: Summary report ───────────────────────────────────────────────────────

def _tab_summary():
    st.markdown("##### Generate a structured summary report")
    names = _dataset_names()
    if not names:
        st.info("No datasets available.")
        return

    name  = st.selectbox("Dataset", names, key="rpt_sum_ds")
    df    = _resolve(name)
    if df is None:
        return

    report_title   = st.text_input("Report title",    value=f"Summary Report — {name}", key="rpt_title")
    report_author  = st.text_input("Author / team",   value="DataOps Studio",           key="rpt_author")
    include_nulls  = st.checkbox("Include null analysis",   value=True, key="rpt_nulls")
    include_stats  = st.checkbox("Include descriptive stats", value=True, key="rpt_stats")
    include_schema = st.checkbox("Include schema / types",    value=True, key="rpt_schema")

    if st.button("Build Report Preview", type="primary"):
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        sections = []

        # Header
        sections.append({
            "type": "header",
            "title": report_title,
            "author": report_author,
            "generated": ts,
            "dataset": name,
            "rows": len(df),
            "columns": len(df.columns),
        })

        if include_schema:
            schema_rows = [
                {"Column": c, "Type": str(t), "Null %": f"{df[c].isna().mean()*100:.1f}%", "Unique": df[c].nunique()}
                for c, t in df.dtypes.items()
            ]
            sections.append({"type": "schema", "data": schema_rows})

        if include_nulls:
            null_rows = [
                {"Column": c, "Null count": int(df[c].isna().sum()), "Null %": f"{df[c].isna().mean()*100:.1f}%"}
                for c in df.columns if df[c].isna().any()
            ]
            sections.append({"type": "nulls", "data": null_rows})

        num_cols = df.select_dtypes(include="number").columns.tolist()
        if include_stats and num_cols:
            desc = df[num_cols].describe().T.reset_index().rename(columns={"index": "column"})
            sections.append({"type": "stats", "data": desc.round(4).to_dict(orient="records")})

        # Render preview
        st.markdown(f"## {report_title}")
        st.caption(f"Author: {report_author}  ·  Generated: {ts}  ·  Dataset: {name}")
        st.markdown(f"**{len(df):,} rows × {len(df.columns)} columns**")
        st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

        for sec in sections[1:]:  # skip header dict
            if sec["type"] == "schema":
                st.markdown("### Schema")
                st.dataframe(pd.DataFrame(sec["data"]), use_container_width=True)

            elif sec["type"] == "nulls":
                if sec["data"]:
                    st.markdown("### Null Analysis")
                    st.dataframe(pd.DataFrame(sec["data"]), use_container_width=True)
                else:
                    st.markdown("### Null Analysis")
                    st.success("No null values found.")

            elif sec["type"] == "stats":
                st.markdown("### Descriptive Statistics")
                st.dataframe(pd.DataFrame(sec["data"]), use_container_width=True)

        # Export the report as JSON
        json_report = json.dumps(
            [{**s, "data": s["data"] if isinstance(s.get("data"), list) else []} for s in sections],
            indent=2, default=str
        )
        st.download_button(
            "⬇  Download Report (JSON)",
            json_report.encode(),
            f"report_{name}_{datetime.datetime.now().strftime('%Y%m%d')}.json",
            "application/json",
        )

        # Export full data with stats sheet as Excel
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Data")
            pd.DataFrame([{
                "report_title": report_title,
                "author": report_author,
                "generated": ts,
                "rows": len(df),
                "columns": len(df.columns),
                "null_total": int(df.isna().sum().sum()),
            }]).to_excel(writer, index=False, sheet_name="Report Meta")
            if num_cols:
                df[num_cols].describe().T.round(4).to_excel(writer, sheet_name="Statistics")
            schema_df = pd.DataFrame([{"Column": c, "Type": str(t)} for c, t in df.dtypes.items()])
            schema_df.to_excel(writer, index=False, sheet_name="Schema")

        st.download_button(
            "⬇  Download Report (Excel)",
            buf.getvalue(),
            f"report_{name}_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="rpt_xl_dl",
        )


# ── Tab: Multi-sheet workbook ─────────────────────────────────────────────────

def _tab_workbook():
    st.markdown("##### Build a multi-sheet Excel workbook from multiple datasets")
    names = _dataset_names()
    if not names:
        st.info("No datasets available.")
        return

    selected = st.multiselect("Datasets to include (one sheet each)", names, key="wb_sel")

    sheet_names: dict[str, str] = {}
    for ds in selected:
        default_sheet = ds[:31].replace("[", "").replace("]", "").replace(" ", "_")
        sheet_names[ds] = st.text_input(
            f"Sheet name for **{ds}**",
            value=default_sheet,
            max_chars=31,
            key=f"wb_sheet_{ds}",
        )

    wb_title = st.text_input("Workbook file name", value="workbook", key="wb_name")

    if selected and st.button("Build Workbook", type="primary"):
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            for ds in selected:
                df = _resolve(ds)
                if df is not None:
                    sname = sheet_names.get(ds, ds[:31]) or ds[:31]
                    df.to_excel(writer, index=False, sheet_name=sname)
        ts    = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        fname = f"{wb_title.strip() or 'workbook'}_{ts}.xlsx"
        st.download_button(
            f"⬇  Download {fname}",
            buf.getvalue(),
            fname,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="wb_dl",
        )
        st.success(f"Workbook ready with {len(selected)} sheet(s).")


# ── Main render ───────────────────────────────────────────────────────────────

def render():
    page_header(
        "Report Generator",
        "Export data, build structured summary reports, and assemble multi-sheet workbooks.",
    )
    tabs = st.tabs(["⬇  Export", "📄  Summary Report", "📒  Workbook Builder"])
    with tabs[0]: _tab_export()
    with tabs[1]: _tab_summary()
    with tabs[2]: _tab_workbook()
