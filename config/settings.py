"""
config/settings.py — App configuration and global CSS theme
"""

import streamlit as st

APP_CONFIG = {
    "title":   "Builder App",
    "icon":    "⬡",
    "version": "1.0.0",
}

CSS_FILES = [
    "base.css",
    "sidebar.css",
    "cards.css",
    "tables.css",
    "forms.css",
]


def apply_global_styles() -> None:
    """Inject the global CSS into the Streamlit app."""
    st.markdown(CSS_FILES, unsafe_allow_html=True)


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
