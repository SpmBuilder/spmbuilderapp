"""
config/settings.py — App configuration and global CSS theme
"""

from pathlib import Path
import streamlit as st

APP_CONFIG = {
    "title":   "SBC Builder",
    "icon":    "⬡",
    "version": "1.0.0",
}

# CSS files live in config/css/, loaded in this order (later files can
# override earlier ones, so keep base.css first).
CSS_DIR = Path(__file__).parent / "css"

CSS_FILES = [
    "base.css",
    "sidebar.css",
    "cards.css",
    "tables.css",
    "forms.css",
]


def apply_global_styles() -> None:
    """
    Inject the global CSS into the Streamlit app.

    NOTE: this previously did `st.markdown(CSS_FILES, unsafe_allow_html=True)`,
    which just printed the list of filenames as text — none of the actual
    CSS was ever applied. This reads each file from disk and wraps the
    combined result in a single <style> tag.
    """
    css_chunks = []
    for filename in CSS_FILES:
        path = CSS_DIR / filename
        if path.exists():
            css_chunks.append(path.read_text())
        else:
            # Fail loudly in the UI rather than silently skipping styles,
            # so a missing/renamed file doesn't go unnoticed again.
            st.warning(f"Stylesheet not found: {path}")

    combined_css = "\n".join(css_chunks)
    st.markdown(f"<style>{combined_css}</style>", unsafe_allow_html=True)


def page_header(title: str, subtitle: str = "") -> None:
    """Render a consistent page header."""
    st.markdown(
        f"""
        <div class="page-header">
            <h1>{title}</h1>
            {"<p>" + subtitle + "</p>" if subtitle else ""}
        </div>
        """,
        unsafe_allow_html=True,
    )


def metric_card(label: str, value: str, sub: str = "") -> None:
    """Render a single metric card."""
    st.markdown(
        f"""
        <div class="card">
            <div class="card-title">{label}</div>
            <div class="card-value">{value}</div>
            {"<div class='card-sub'>" + sub + "</div>" if sub else ""}
        </div>
        """,
        unsafe_allow_html=True,
    )


def badge(text: str, variant: str = "dark") -> str:
    """Return an HTML badge string."""
    return f'<span class="badge badge-{variant}">{text}</span>'