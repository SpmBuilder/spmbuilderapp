import streamlit as st
from components.sidebar import render_sidebar

st.set_page_config(
    page_title="BUILDER Platform",
    page_icon= "images/image.png",
    layout="wide"
)

render_sidebar()

page = st.session_state.page


# =========================
# ROUTER
# =========================
if page == "dashboard":
    from modules.dashboard import show_dashboard
    show_dashboard()

elif page == "data_sources":
    from modules.data_sources import show_data_sources
    show_data_sources()

elif page == "repository":
    from modules.repository import show_repository
    show_repository()

elif page == "validation":
    from modules.validation import show_validation
    show_validation()

elif page == "merge":
    from modules.merge import show_merge
    show_merge()

elif page == "query_builder":
    from modules.query_builder import show_query_builder
    show_query_builder()

elif page == "analytics":
    from modules.analytics import show_analytics
    show_analytics()

elif page == "templates":
    from modules.templates import show_templates
    show_templates()

elif page == "reports":
    from modules.reports import show_reports
    show_reports()

elif page == "scheduler":
    from modules.scheduler import show_scheduler
    show_scheduler()

elif page == "users":
    from modules.users import show_users
    show_users()

elif page == "settings":
    from modules.settings import show_settings
    show_settings()

import streamlit as st

st.markdown(
    """
    <style>

    /* Global background */
    .stApp {
        background-color: #F7F9FC;
    }

    /* Sidebar styling */
    section[data-testid="stSidebar"] {
        background-color: #0B1F3B;
    }

    /* Sidebar text */
    section[data-testid="stSidebar"] * {
        color: white;
    }

    /* Metric cards */
    div[data-testid="metric-container"] {
        background-color: #FFFFFF;
        border: 1px solid #E5E7EB;
        padding: 12px;
        border-radius: 10px;
        box-shadow: 0px 1px 3px rgba(0,0,0,0.05);
    }

    /* Buttons */
    .stButton > button {
        background-color: #1E5EFF;
        color: white;
        border-radius: 8px;
        border: none;
    }

    .stButton > button:hover {
        background-color: #1749CC;
        color: white;
    }

    /* Dataframes */
    .stDataFrame {
        border-radius: 10px;
    }

    </style>
    """,
    unsafe_allow_html=True
)