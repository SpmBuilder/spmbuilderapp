import streamlit as st
import base64


# =========================
# INIT NAV STATE
# =========================
def init_navigation():

    if "page" not in st.session_state:
        st.session_state.page = "dashboard"


# =========================
# NAV HANDLER
# =========================
def navigate(page_name: str):
    st.session_state.page = page_name


def get_base64(img_path):
    with open(img_path, "rb") as f:
        return base64.b64encode(f.read()).decode()

img_base64 = get_base64("images/image.png")



# =========================
# SIDEBAR UI
# =========================
def render_sidebar():

    init_navigation()

    with st.sidebar:

        # ================= HEADER =================
        st.markdown(
            """
            <div style="text-align:center; padding:12px;">
                <div style="font-size:20px; font-weight:600; color:#FFFFFF;">
                    BUILDER
                </div>
                <div style="font-size:11px; color:#9CC2FF; letter-spacing:1px;">
                    DATA • ANALYTICS • AUTOMATION
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        # ================= DASHBOARD =================
        st.button(
            "📊 Dashboard",
            use_container_width=True,
            on_click=navigate,
            args=("dashboard",)
        )

        st.markdown("---")

        # ================= DATA MANAGEMENT =================
        with st.expander("📁 Data Management", expanded=True):

            st.button("☁ Data Sources", use_container_width=True,
                      on_click=navigate, args=("data_sources",))

            st.button("🗄 Data Repository", use_container_width=True,
                      on_click=navigate, args=("repository",))

            st.button("✅ Data Validation", use_container_width=True,
                      on_click=navigate, args=("validation",))

            st.button("🔗 Data Merge", use_container_width=True,
                      on_click=navigate, args=("merge",))

        # ================= ANALYTICS =================
        with st.expander("📈 Analytics", expanded=False):

            st.button("🔍 Query Builder", use_container_width=True,
                      on_click=navigate, args=("query_builder",))

            st.button("📊 Analytics", use_container_width=True,
                      on_click=navigate, args=("analytics",))

        # ================= REPORTING =================
        with st.expander("📑 Reporting", expanded=False):

            st.button("🧩 Templates", use_container_width=True,
                      on_click=navigate, args=("templates",))

            st.button("📄 Reports", use_container_width=True,
                      on_click=navigate, args=("reports",))

            st.button("⏱ Scheduler", use_container_width=True,
                      on_click=navigate, args=("scheduler",))

        # ================= ADMIN =================
        with st.expander("⚙ Administration", expanded=False):

            st.button("👤 Users", use_container_width=True,
                      on_click=navigate, args=("users",))

            st.button("🔐 Settings", use_container_width=True,
                      on_click=navigate, args=("settings",))

        st.markdown("---")

        # ================= FOOTER =================
        st.markdown(
            """
            <div style="font-size:12px; color:gray;">
                <b>User:</b> Administrator<br>
                <b>Version:</b> 1.0.0<br>
                <b>Status:</b> 🟢 Online
            </div>
            """,
            unsafe_allow_html=True
        )