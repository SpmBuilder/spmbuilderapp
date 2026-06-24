"""
modules/data_validator.py — Schema, type, null, range, and custom rule validation
"""

import re
import streamlit as st
import pandas as pd
from config.settings import page_header, badge

# ── Helpers ───────────────────────────────────────────────────────────────────

def _get_datasets() -> dict:
    return st.session_state.get("uploaded_datasets", {})


def _resolve(name: str) -> pd.DataFrame | None:
    return _get_datasets().get(name, {}).get("df")


def _dataset_names() -> list[str]:
    return list(_get_datasets().keys())


# ── Validation rules ──────────────────────────────────────────────────────────

def _check_nulls(df: pd.DataFrame, threshold: float) -> list[dict]:
    issues = []
    for col in df.columns:
        pct = df[col].isna().mean() * 100
        if pct > threshold:
            issues.append({
                "Column": col, "Rule": "Null threshold",
                "Severity": "Error" if pct > 50 else "Warning",
                "Detail": f"{pct:.1f}% null (threshold {threshold}%)",
                "Affected Rows": int(df[col].isna().sum()),
            })
    return issues


def _check_duplicates(df: pd.DataFrame, keys: list[str]) -> list[dict]:
    issues = []
    if not keys:
        dup_count = df.duplicated().sum()
        if dup_count:
            issues.append({
                "Column": "(all)", "Rule": "Duplicate rows",
                "Severity": "Warning",
                "Detail": f"{dup_count:,} fully duplicate rows found.",
                "Affected Rows": int(dup_count),
            })
    else:
        valid_keys = [k for k in keys if k in df.columns]
        if valid_keys:
            dup_count = df.duplicated(subset=valid_keys).sum()
            if dup_count:
                issues.append({
                    "Column": ", ".join(valid_keys), "Rule": "Duplicate keys",
                    "Severity": "Error",
                    "Detail": f"{dup_count:,} duplicate key combinations.",
                    "Affected Rows": int(dup_count),
                })
    return issues


def _check_type(df: pd.DataFrame, col: str, expected_type: str) -> list[dict]:
    issues = []
    s = df[col].dropna()
    if expected_type == "numeric":
        bad = pd.to_numeric(s, errors="coerce").isna().sum()
        if bad:
            issues.append({
                "Column": col, "Rule": "Type: numeric",
                "Severity": "Error",
                "Detail": f"{bad:,} non-numeric values.",
                "Affected Rows": int(bad),
            })
    elif expected_type == "date":
        bad = pd.to_datetime(s, errors="coerce").isna().sum()
        if bad:
            issues.append({
                "Column": col, "Rule": "Type: date",
                "Severity": "Error",
                "Detail": f"{bad:,} unparseable date values.",
                "Affected Rows": int(bad),
            })
    elif expected_type == "email":
        pattern = r"^[\w\.\+\-]+@[\w\-]+\.[a-z]{2,}$"
        bad = (~s.astype(str).str.match(pattern, case=False)).sum()
        if bad:
            issues.append({
                "Column": col, "Rule": "Type: email",
                "Severity": "Warning",
                "Detail": f"{bad:,} values fail email format.",
                "Affected Rows": int(bad),
            })
    elif expected_type == "non-empty string":
        bad = (s.astype(str).str.strip() == "").sum()
        if bad:
            issues.append({
                "Column": col, "Rule": "Non-empty string",
                "Severity": "Warning",
                "Detail": f"{bad:,} empty/whitespace-only strings.",
                "Affected Rows": int(bad),
            })
    return issues


def _check_range(df: pd.DataFrame, col: str, min_val, max_val) -> list[dict]:
    issues = []
    numeric = pd.to_numeric(df[col], errors="coerce")
    if min_val is not None:
        under = (numeric < min_val).sum()
        if under:
            issues.append({
                "Column": col, "Rule": f"Range min ({min_val})",
                "Severity": "Error",
                "Detail": f"{under:,} values below minimum.",
                "Affected Rows": int(under),
            })
    if max_val is not None:
        over = (numeric > max_val).sum()
        if over:
            issues.append({
                "Column": col, "Rule": f"Range max ({max_val})",
                "Severity": "Error",
                "Detail": f"{over:,} values above maximum.",
                "Affected Rows": int(over),
            })
    return issues


def _check_regex(df: pd.DataFrame, col: str, pattern: str) -> list[dict]:
    issues = []
    try:
        compiled = re.compile(pattern)
        bad = (~df[col].astype(str).str.match(compiled)).sum()
        if bad:
            issues.append({
                "Column": col, "Rule": f"Regex: {pattern}",
                "Severity": "Warning",
                "Detail": f"{bad:,} values don't match pattern.",
                "Affected Rows": int(bad),
            })
    except re.error as e:
        issues.append({
            "Column": col, "Rule": "Regex",
            "Severity": "Warning",
            "Detail": f"Invalid regex: {e}",
            "Affected Rows": 0,
        })
    return issues


# ── Severity badge HTML ───────────────────────────────────────────────────────

def _sev_badge(sev: str) -> str:
    color = "badge-red" if sev == "Error" else "badge-yellow"
    return f'<span class="badge {color}">{sev}</span>'


# ── Tab: Quick validation ──────────────────────────────────────────────────────

def _tab_quick():
    st.markdown("##### Automated dataset health check")
    names = _dataset_names()
    if not names:
        st.info("No datasets in the workspace. Upload data first.")
        return

    name = st.selectbox("Dataset", names, key="val_quick_ds")
    df = _resolve(name)
    if df is None:
        return

    c1, c2, c3 = st.columns(3)
    c1.metric("Rows", f"{len(df):,}")
    c2.metric("Columns", len(df.columns))
    c3.metric("Total Nulls", f"{df.isna().sum().sum():,}")

    null_thresh = st.slider("Null % alert threshold", 0, 100, 20, key="val_null_thresh")
    dup_check   = st.checkbox("Check for duplicate rows", value=True, key="val_dup")

    if st.button("Run Health Check", type="primary"):
        all_issues = []
        all_issues.extend(_check_nulls(df, null_thresh))
        if dup_check:
            all_issues.extend(_check_duplicates(df, []))

        if not all_issues:
            st.success("✅ No issues detected.")
        else:
            errors   = [i for i in all_issues if i["Severity"] == "Error"]
            warnings = [i for i in all_issues if i["Severity"] == "Warning"]

            st.markdown(
                f"""
                <div style="display:flex;gap:1rem;margin-bottom:1rem;">
                    <div class="card" style="flex:1;">
                        <div class="card-title">Errors</div>
                        <div class="card-value" style="color:#f44336;">{len(errors)}</div>
                    </div>
                    <div class="card" style="flex:1;">
                        <div class="card-title">Warnings</div>
                        <div class="card-value" style="color:#ffb300;">{len(warnings)}</div>
                    </div>
                    <div class="card" style="flex:1;">
                        <div class="card-title">Total Issues</div>
                        <div class="card-value">{len(all_issues)}</div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            result_df = pd.DataFrame(all_issues)
            st.dataframe(result_df, use_container_width=True)
            st.session_state["val_last_report"] = result_df


# ── Tab: Custom rules ─────────────────────────────────────────────────────────

def _tab_custom():
    st.markdown("##### Define per-column validation rules")
    names = _dataset_names()
    if not names:
        st.info("No datasets in the workspace.")
        return

    name = st.selectbox("Dataset", names, key="val_custom_ds")
    df = _resolve(name)
    if df is None:
        return

    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
    rules: list[dict] = st.session_state.setdefault("custom_rules", [])

    with st.expander("➕  Add a rule", expanded=True):
        col_sel  = st.selectbox("Column", df.columns.tolist(), key="cr_col")
        rule_type = st.selectbox(
            "Rule type",
            ["Not null", "Numeric type", "Date type", "Email format",
             "Non-empty string", "Min value", "Max value", "Regex pattern"],
            key="cr_type",
        )
        rule_param = None
        if rule_type in ("Min value", "Max value"):
            rule_param = st.number_input("Threshold", key="cr_param")
        elif rule_type == "Regex pattern":
            rule_param = st.text_input("Pattern", placeholder=r"^\d{4}-\d{2}-\d{2}$", key="cr_param_r")

        if st.button("Add Rule", key="add_rule"):
            rules.append({"col": col_sel, "type": rule_type, "param": rule_param})
            st.success(f"Rule added: {col_sel} → {rule_type}")

    if rules:
        st.markdown(f"**Active rules ({len(rules)})**")
        for i, r in enumerate(rules):
            c1, c2 = st.columns([5, 1])
            c1.markdown(
                f'`{r["col"]}` — {r["type"]}'
                + (f' `{r["param"]}`' if r["param"] is not None else "")
            )
            if c2.button("✕", key=f"rm_rule_{i}"):
                rules.pop(i)
                st.rerun()

        if st.button("Run Custom Validation", type="primary"):
            all_issues = []
            for r in rules:
                col, rtype, param = r["col"], r["type"], r["param"]
                if col not in df.columns:
                    continue
                if rtype == "Not null":
                    all_issues.extend(_check_nulls(df[[col]], 0))
                elif rtype == "Numeric type":
                    all_issues.extend(_check_type(df, col, "numeric"))
                elif rtype == "Date type":
                    all_issues.extend(_check_type(df, col, "date"))
                elif rtype == "Email format":
                    all_issues.extend(_check_type(df, col, "email"))
                elif rtype == "Non-empty string":
                    all_issues.extend(_check_type(df, col, "non-empty string"))
                elif rtype == "Min value":
                    all_issues.extend(_check_range(df, col, param, None))
                elif rtype == "Max value":
                    all_issues.extend(_check_range(df, col, None, param))
                elif rtype == "Regex pattern" and param:
                    all_issues.extend(_check_regex(df, col, param))

            if not all_issues:
                st.success("✅ All custom rules passed.")
            else:
                st.error(f"{len(all_issues)} rule violation(s) found.")
                st.dataframe(pd.DataFrame(all_issues), use_container_width=True)


# ── Tab: Schema comparison ────────────────────────────────────────────────────

def _tab_schema():
    st.markdown("##### Compare schemas of two datasets")
    names = _dataset_names()
    if len(names) < 2:
        st.info("Need at least 2 datasets.")
        return

    c1, c2 = st.columns(2)
    a_name = c1.selectbox("Dataset A", names, key="sch_a")
    b_name = c2.selectbox("Dataset B", names, index=1, key="sch_b")

    df_a = _resolve(a_name)
    df_b = _resolve(b_name)
    if df_a is None or df_b is None:
        return

    cols_a = {c: str(t) for c, t in df_a.dtypes.items()}
    cols_b = {c: str(t) for c, t in df_b.dtypes.items()}
    all_cols = sorted(set(cols_a) | set(cols_b))

    rows = []
    for col in all_cols:
        in_a = col in cols_a
        in_b = col in cols_b
        type_match = cols_a.get(col) == cols_b.get(col) if (in_a and in_b) else None
        rows.append({
            "Column":    col,
            "In A":      "✅" if in_a else "—",
            "Type A":    cols_a.get(col, "—"),
            "In B":      "✅" if in_b else "—",
            "Type B":    cols_b.get(col, "—"),
            "Match":     "✅" if type_match else ("⚠️" if type_match is False else "❌"),
        })

    schema_df = pd.DataFrame(rows)
    matches    = schema_df["Match"].eq("✅").sum()
    mismatches = len(schema_df) - matches

    c3, c4, c5 = st.columns(3)
    c3.metric("Total Columns", len(schema_df))
    c4.metric("Matched",    matches)
    c5.metric("Mismatches", mismatches)

    st.dataframe(schema_df, use_container_width=True)


# ── Main render ───────────────────────────────────────────────────────────────

def render():
    page_header(
        "Data Validator",
        "Run health checks, custom rules, and schema comparisons on your datasets.",
    )
    tabs = st.tabs(["🩺  Health Check", "📋  Custom Rules", "🔍  Schema Compare"])
    with tabs[0]: _tab_quick()
    with tabs[1]: _tab_custom()
    with tabs[2]: _tab_schema()
