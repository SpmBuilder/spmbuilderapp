"""
utils/helpers.py — Shared utility functions
"""

import re
import hashlib
import datetime
import pandas as pd
import streamlit as st


# ── Data utilities ────────────────────────────────────────────────────────────

def slugify(text: str) -> str:
    """Convert text to a safe identifier."""
    return re.sub(r"[^\w]", "_", text.strip().lower())


def human_size(n_bytes: int) -> str:
    """Format byte count as human-readable string."""
    for unit in ("B", "KB", "MB", "GB"):
        if n_bytes < 1024:
            return f"{n_bytes:.1f} {unit}"
        n_bytes /= 1024
    return f"{n_bytes:.1f} TB"


def df_fingerprint(df: pd.DataFrame) -> str:
    """Return a short hash that changes when df shape or content changes."""
    raw = f"{df.shape}{list(df.columns)}{df.head(10).to_csv()}"
    return hashlib.md5(raw.encode()).hexdigest()[:8]


def now_label() -> str:
    """ISO timestamp string for file names / report metadata."""
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def safe_cast_numeric(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    """Attempt coercing listed columns to numeric; leave untouched on failure."""
    df = df.copy()
    for col in columns:
        if col in df.columns:
            coerced = pd.to_numeric(df[col], errors="coerce")
            if coerced.notna().sum() > 0:
                df[col] = coerced
    return df


def infer_date_columns(df: pd.DataFrame, threshold: float = 0.80) -> list[str]:
    """Return object columns that look like dates (>= threshold parse rate)."""
    candidates = []
    for col in df.select_dtypes(include="object").columns:
        sample = df[col].dropna().head(200)
        if sample.empty:
            continue
        parsed = pd.to_datetime(sample, errors="coerce")
        if parsed.notna().mean() >= threshold:
            candidates.append(col)
    return candidates


def truncate_string_columns(df: pd.DataFrame, max_len: int = 200) -> pd.DataFrame:
    """Truncate long string columns for display purposes."""
    df = df.copy()
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).str[:max_len]
    return df


# ── Session state utilities ───────────────────────────────────────────────────

def get_workspace_datasets() -> dict:
    return st.session_state.get("uploaded_datasets", {})


def add_to_workspace(name: str, df: pd.DataFrame, source: str = "unknown"):
    store = st.session_state.setdefault("uploaded_datasets", {})
    store[name] = {
        "df": df,
        "source": source,
        "rows": len(df),
        "cols": len(df.columns),
        "added": now_label(),
    }


def remove_from_workspace(name: str):
    store = st.session_state.get("uploaded_datasets", {})
    store.pop(name, None)


def list_workspace_names() -> list[str]:
    return list(get_workspace_datasets().keys())


# ── Streamlit convenience wrappers ────────────────────────────────────────────

def confirm_button(label: str, key: str, danger: bool = False) -> bool:
    """
    Two-click confirmation button.
    First click sets a flag; second click returns True.
    """
    flag_key = f"_confirm_{key}"
    if not st.session_state.get(flag_key):
        if st.button(label, key=key):
            st.session_state[flag_key] = True
            st.rerun()
        return False
    else:
        col1, col2 = st.columns(2)
        confirmed = col1.button("✓ Confirm", key=f"{key}_yes", type="primary")
        cancelled = col2.button("✗ Cancel",  key=f"{key}_no")
        if confirmed:
            st.session_state.pop(flag_key)
            return True
        if cancelled:
            st.session_state.pop(flag_key)
            st.rerun()
        return False


def empty_state(message: str, icon: str = "📭"):
    """Render a centered empty-state message."""
    st.markdown(
        f"""
        <div style="text-align:center;padding:3rem 1rem;color:#444;">
            <div style="font-size:2.5rem;margin-bottom:0.75rem;">{icon}</div>
            <div style="font-size:0.88rem;">{message}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
