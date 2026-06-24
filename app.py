"""
BUILDER APP — Main Entry Point
"""
from pathlib import Path
import streamlit as st
from config.settings import APP_CONFIG, apply_global_styles
from modules.db_connector import render as db_render
from modules.data_uploader import render as upload_render
from modules.data_merger import render as merger_render
from modules.data_validator import render as validator_render
from modules.data_analytics import render as analytics_render
from modules.template_mapper import render as template_render
from modules.report_generator import render as report_render

st.set_page_config(
    page_title=APP_CONFIG["title"],
    page_icon=APP_CONFIG["icon"],
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        "Get Help": None,
        "Report a bug": None,
        "About": f"**{APP_CONFIG['title']}** v{APP_CONFIG['version']}",
    },
)

apply_global_styles()


# ── Sidebar navigation ────────────────────────────────────────────────────────
def render_sidebar() -> str:

    with st.sidebar:

        # =========================
        # BRAND
        # =========================

        st.markdown(
            """
            <div class="sidebar-brand">
                <div class="brand-title">
                    BUILDER
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.divider()

        # =========================
        # NAVIGATION
        # =========================

        nav_items = {
            "🗄️ Database":      "Database",
            "📂 Data Upload":   "Data Upload",
            "🔗 Data Merger":   "Data Merger",
            "✅ Validator":     "Validator",
            "📊 Analytics":     "Analytics",
            "🗂️ Template Map":  "Template Map",
            "📄 Reports":       "Reports",
        }

        if "active_page" not in st.session_state:
            st.session_state.active_page = "Database"

        for label, page in nav_items.items():

            active = st.session_state.active_page == page

            if active:
                st.markdown(
                    '<div class="nav-btn-active">',
                    unsafe_allow_html=True,
                )

            clicked = st.button(
                label,
                key=f"nav_{page}",
                use_container_width=True,
            )

            if active:
                st.markdown("</div>", unsafe_allow_html=True)

            if clicked:
                st.session_state.active_page = page
                st.rerun()

        st.divider()

        # =========================
        # STATUS
        # =========================

        db_ok = st.session_state.get(
            "db_connected",
            False
        )

        status_class = (
            "badge-active"
            if db_ok
            else "badge-inactive"
        )

        status_text = (
            "DATABASE CONNECTED"
            if db_ok
            else "NO DATABASE"
        )

        st.markdown(
            f"""
            <div style="margin-top:12px;">
                <span class="badge {status_class}">
                    {status_text}
                </span>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown(
            f"""
            <div style="
                margin-top:20px;
                color:#71717A;
                font-size:0.75rem;
                text-align:center;
            ">
                Version {APP_CONFIG["version"]}
            </div>
            """,
            unsafe_allow_html=True,
        )

    return st.session_state.active_page


# ── Page dispatcher ───────────────────────────────────────────────────────────

def main() -> None:
    page = render_sidebar()

    dispatch = {
        "Database":     db_render,
        "Data Upload":  upload_render,
        "Data Merger":  merger_render,
        "Validator":    validator_render,
        "Analytics":    analytics_render,
        "Template Map": template_render,
        "Reports":      report_render,
    }

    renderer = dispatch.get(page)
    if renderer:
        renderer()
    else:
        st.error("Page not found.")


if __name__ == "__main__":
    main()
