"""
modules/template_mapper.py — Map dataset columns onto named template fields
"""

import io
import json
import streamlit as st
import pandas as pd
from config.settings import page_header

# ── Helpers ───────────────────────────────────────────────────────────────────

def _get_datasets() -> dict:
    return st.session_state.get("uploaded_datasets", {})

def _dataset_names() -> list[str]:
    return list(_get_datasets().keys())

def _resolve(name: str) -> pd.DataFrame | None:
    return _get_datasets().get(name, {}).get("df")

def _store(name: str, df: pd.DataFrame, source: str):
    store = st.session_state.setdefault("uploaded_datasets", {})
    store[name] = {"df": df, "source": source, "rows": len(df), "cols": len(df.columns)}


# ── Built-in template definitions ─────────────────────────────────────────────

BUILTIN_TEMPLATES = {
    "Sales Report": {
        "description": "Standard sales summary with revenue and customer fields.",
        "fields": ["date", "customer_name", "product", "quantity", "unit_price", "total_revenue", "region", "sales_rep"],
    },
    "HR Employee": {
        "description": "Employee master record template.",
        "fields": ["employee_id", "first_name", "last_name", "department", "job_title", "hire_date", "salary", "email", "manager"],
    },
    "Financial Ledger": {
        "description": "General ledger entry format.",
        "fields": ["entry_date", "account_code", "account_name", "debit", "credit", "balance", "reference", "description"],
    },
    "Customer Master": {
        "description": "CRM customer record.",
        "fields": ["customer_id", "company_name", "contact_name", "email", "phone", "country", "segment", "created_date"],
    },
    "Inventory": {
        "description": "Warehouse/inventory tracking.",
        "fields": ["sku", "product_name", "category", "quantity_on_hand", "reorder_point", "unit_cost", "location", "last_updated"],
    },
}


# ── Tab: Template builder ─────────────────────────────────────────────────────

def _tab_builder():
    st.markdown("##### Define or edit a template")

    templates: dict = st.session_state.setdefault("custom_templates", {})

    c1, c2 = st.columns([2, 1])
    tpl_name = c1.text_input("Template name", placeholder="My Template", key="tb_name")
    tpl_desc = c2.text_input("Description",   placeholder="Optional description", key="tb_desc")

    st.markdown("**Fields** — one per line")
    raw_fields = st.text_area(
        "Fields",
        height=160,
        placeholder="date\ncustomer_name\nproduct\nquantity\nunit_price",
        label_visibility="collapsed",
        key="tb_fields",
    )

    col_save, col_clear = st.columns([1, 1])

    with col_save:
        if st.button("Save Template", type="primary", use_container_width=True):
            if not tpl_name.strip():
                st.warning("Template name is required.")
            elif not raw_fields.strip():
                st.warning("Enter at least one field.")
            else:
                fields = [f.strip() for f in raw_fields.splitlines() if f.strip()]
                templates[tpl_name.strip()] = {
                    "description": tpl_desc.strip(),
                    "fields": fields,
                }
                st.success(f"Template **{tpl_name}** saved with {len(fields)} fields.")

    with col_clear:
        if st.button("Load Built-in →", use_container_width=True):
            st.session_state["tb_show_builtins"] = True

    if st.session_state.get("tb_show_builtins"):
        chosen_builtin = st.selectbox("Choose built-in", list(BUILTIN_TEMPLATES.keys()), key="tb_bi_sel")
        if st.button("Load", key="tb_bi_load", type="primary"):
            bt = BUILTIN_TEMPLATES[chosen_builtin]
            templates[chosen_builtin] = bt
            st.success(f"Loaded **{chosen_builtin}**.")
            st.session_state["tb_show_builtins"] = False

    # List saved templates
    all_tpls = {**BUILTIN_TEMPLATES, **templates}
    if all_tpls:
        st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
        st.markdown("**Saved templates**")
        for tname, tdata in all_tpls.items():
            builtin_tag = " *(built-in)*" if tname in BUILTIN_TEMPLATES else ""
            with st.expander(f"🗂️  {tname}{builtin_tag}  —  {len(tdata['fields'])} fields"):
                st.caption(tdata.get("description", ""))
                st.code(", ".join(tdata["fields"]), language="text")
                if tname not in BUILTIN_TEMPLATES:
                    if st.button("Delete", key=f"del_tpl_{tname}"):
                        del templates[tname]
                        st.rerun()

    # Export
    if st.button("Export all templates (JSON)"):
        payload = json.dumps(all_tpls, indent=2)
        st.download_button(
            "⬇ Download templates.json",
            payload.encode(),
            "templates.json",
            "application/json",
        )


# ── Tab: Column mapper ────────────────────────────────────────────────────────

def _tab_mapper():
    st.markdown("##### Map dataset columns to a template")

    ds_names = _dataset_names()
    if not ds_names:
        st.info("No datasets in the workspace. Upload data first.")
        return

    custom_tpls  = st.session_state.get("custom_templates", {})
    all_templates = {**BUILTIN_TEMPLATES, **custom_tpls}

    if not all_templates:
        st.info("No templates defined. Create one in the Template Builder tab.")
        return

    c1, c2 = st.columns(2)
    ds_name  = c1.selectbox("Source dataset", ds_names, key="map_ds")
    tpl_name = c2.selectbox("Template",       list(all_templates.keys()), key="map_tpl")

    df = _resolve(ds_name)
    if df is None:
        return

    template = all_templates[tpl_name]
    tpl_fields  = template["fields"]
    ds_columns  = ["(skip)"] + df.columns.tolist()

    st.markdown(f"**Mapping** — template: `{tpl_name}` · dataset: `{ds_name}`")
    st.caption(template.get("description", ""))

    mapping: dict[str, str] = {}

    for field in tpl_fields:
        # Auto-suggest: exact match, then case-insensitive
        default_idx = 0
        for i, col in enumerate(ds_columns):
            if col.lower() == field.lower() or col.lower().replace(" ", "_") == field.lower():
                default_idx = i
                break

        chosen = st.selectbox(
            f"**{field}**",
            ds_columns,
            index=default_idx,
            key=f"map_{field}",
        )
        if chosen != "(skip)":
            mapping[field] = chosen

    out_name = st.text_input("Output dataset name", value=f"{ds_name}_mapped", key="map_out")

    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
    c_apply, c_preview = st.columns([1, 1])

    if c_apply.button("Apply Mapping", type="primary", use_container_width=True):
        if not mapping:
            st.warning("Map at least one field.")
        else:
            result = df[list(mapping.values())].rename(columns={v: k for k, v in mapping.items()})
            _store(out_name.strip() or "mapped_output", result, f"mapped:{ds_name}→{tpl_name}")
            st.success(f"Mapped dataset saved as **{out_name}** — {len(result.columns)} columns.")
            st.session_state["last_mapped"] = result
            st.dataframe(result.head(30), use_container_width=True)

    if c_preview.button("Preview mapping", use_container_width=True):
        if mapping:
            preview_rows = []
            for tfield, dcol in mapping.items():
                sample = df[dcol].dropna().iloc[0] if df[dcol].notna().any() else "—"
                preview_rows.append({
                    "Template field": tfield,
                    "Dataset column": dcol,
                    "Sample value":   str(sample),
                })
            st.dataframe(pd.DataFrame(preview_rows), use_container_width=True)

    # Unmapped fields
    unmapped = [f for f in tpl_fields if f not in mapping]
    if unmapped:
        st.warning(f"**{len(unmapped)} field(s) unmapped:** {', '.join(unmapped)}")


# ── Tab: Import template ──────────────────────────────────────────────────────

def _tab_import():
    st.markdown("##### Import templates from a JSON file")
    uploaded = st.file_uploader("Upload templates.json", type=["json"], key="tpl_import_file")

    if uploaded:
        try:
            data = json.load(uploaded)
            customs = st.session_state.setdefault("custom_templates", {})
            count = 0
            for k, v in data.items():
                if isinstance(v, dict) and "fields" in v:
                    customs[k] = v
                    count += 1
            st.success(f"Imported {count} template(s).")
        except Exception as e:
            st.error(f"Could not parse file: {e}")


# ── Main render ───────────────────────────────────────────────────────────────

def render():
    page_header(
        "Template Mapper",
        "Define reusable field templates and map dataset columns to them.",
    )
    tabs = st.tabs(["🗂️  Template Builder", "🔗  Column Mapper", "📥  Import"])
    with tabs[0]: _tab_builder()
    with tabs[1]: _tab_mapper()
    with tabs[2]: _tab_import()
