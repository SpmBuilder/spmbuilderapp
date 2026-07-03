"""
BUILDER APP — Main Entry Point
"""
from pathlib import Path
import streamlit as st
import time
import math
from streamlit.components.v1 import html
from config.settings import APP_CONFIG, apply_global_styles
from modules.data_uploader import render as upload_render
from modules.data_merger import render as merger_render
from modules.data_analytics import render as analytics_render
from modules.template_mapper import render as template_render
from modules.report_generator import render as report_render

# ── Page configuration ──────────────────────────────────────────────────────
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

# ── Loading Screen ──────────────────────────────────────────────────────────
def show_loading_screen():
    """Display the loading screen with 5 stars animation.

    Stars alternate teal / blue to echo Security Bank's two-tone
    "Perfect Harmony" mark instead of a single flat color.
    """

    stars_html = ""
    num_stars = 5
    center_x = 120
    center_y = 120
    radius = 85
    star_colors = ["#00A99D", "#0072BC"]  # teal, blue — alternating

    for i in range(num_stars):
        angle = (i / num_stars) * 2 * math.pi - math.pi / 2
        x = center_x + radius * math.cos(angle)
        y = center_y + radius * math.sin(angle)
        delay = i * 0.2
        color = star_colors[i % 2]
        stars_html += f"""
            <div class="star" style="top: {y-21}px; left: {x-21}px; animation-delay: {delay}s; color: {color};" id="star{i}">
                ★
            </div>
        """

    loading_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            body {{
                margin: 0;
                padding: 0;
                overflow: hidden;
                background: #000000;
            }}

            .loading-container {{
                position: fixed;
                top: 0;
                left: 0;
                width: 100vw;
                height: 100vh;
                display: flex;
                justify-content: center;
                align-items: center;
                background: #000000;
                z-index: 9999;
                flex-direction: column;
            }}

            .stars-container {{
                position: relative;
                width: 240px;
                height: 240px;
                margin-bottom: 30px;
            }}

            .star {{
                position: absolute;
                font-size: 42px;
                text-shadow: 0 0 20px currentColor;
                transition: all 0.3s ease;
                opacity: 0.85;
            }}

            .star-orbit {{
                position: absolute;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                animation: rotate 20s linear infinite;
            }}

            .center-text {{
                position: absolute;
                top: 50%;
                left: 50%;
                transform: translate(-50%, -50%);
                text-align: center;
                font-family: 'Poppins', 'Arial Black', sans-serif;
                z-index: 10;
                width: 130px;
            }}

            .center-text-line1 {{
                font-size: 15px;
                font-weight: 700;
                color: #00A99D;
                letter-spacing: 5px;
                text-shadow: 0 0 20px rgba(0, 169, 157, 0.35);
                margin-bottom: 2px;
            }}

            .center-text-line2 {{
                font-size: 24px;
                font-weight: 700;
                color: #FAFAFA;
                letter-spacing: 1px;
                text-shadow: 0 0 30px rgba(0, 114, 188, 0.25);
            }}

            .sparkle {{
                animation: sparkle 0.5s ease-in-out infinite alternate;
            }}

            .loading-text {{
                color: #9CA3AF;
                font-size: 16px;
                margin-top: 20px;
                font-family: 'Inter', 'Arial', sans-serif;
                letter-spacing: 2px;
                font-weight: 300;
            }}

            @keyframes sparkle {{
                0% {{
                    text-shadow: 0 0 10px currentColor;
                    opacity: 0.85;
                    transform: scale(1);
                }}
                100% {{
                    text-shadow: 0 0 40px currentColor, 0 0 80px currentColor;
                    opacity: 1;
                    transform: scale(1.2);
                }}
            }}

            @keyframes rotate {{
                0% {{ transform: rotate(0deg); }}
                100% {{ transform: rotate(360deg); }}
            }}

            .progress-bar {{
                width: 300px;
                height: 3px;
                background: rgba(255, 255, 255, 0.1);
                border-radius: 2px;
                margin-top: 20px;
                overflow: hidden;
            }}

            .progress-fill {{
                height: 100%;
                background: linear-gradient(90deg, #00A99D, #0072BC);
                border-radius: 2px;
                transition: width 0.5s ease;
                width: 0%;
            }}

            ::-webkit-scrollbar {{
                display: none;
            }}
        </style>
    </head>
    <body>
        <div id="loadingScreen" class="loading-container">
            <div class="stars-container">
                <div class="star-orbit">
                    {stars_html}
                </div>
                <div class="center-text">
                    <div class="center-text-line1">SBC</div>
                    <div class="center-text-line2">BUILDER</div>
                </div>
            </div>
            <div class="loading-text">Loading SBC Builder...</div>
            <div class="progress-bar">
                <div class="progress-fill" id="progressFill"></div>
            </div>
        </div>

        <script>
            const stars = document.querySelectorAll('.star');
            let progress = 0;
            const progressFill = document.getElementById('progressFill');

            function sparkleStars() {{
                stars.forEach(star => star.classList.remove('sparkle'));
                const numToSparkle = Math.floor(Math.random() * 3) + 1;
                const shuffled = Array.from(stars).sort(() => 0.5 - Math.random());
                for (let i = 0; i < numToSparkle && i < shuffled.length; i++) {{
                    shuffled[i].classList.add('sparkle');
                }}
            }}

            function updateProgress() {{
                if (progress < 100) {{
                    progress += Math.random() * 3 + 1;
                    if (progress > 100) progress = 100;
                    progressFill.style.width = progress + '%';
                    sparkleStars();
                    if (progress < 100) {{
                        setTimeout(updateProgress, 100);
                    }}
                }}
            }}

            setTimeout(updateProgress, 500);
        </script>
    </body>
    </html>
    """

    html(loading_html, height=800, scrolling=False)

# ── Main app ──────────────────────────────────────────────────────────────
def main():
    if "loading_complete" not in st.session_state:
        st.session_state.loading_complete = False

    if not st.session_state.loading_complete:
        show_loading_screen()
        time.sleep(2.5)
        st.session_state.loading_complete = True
        st.rerun()
        return

    apply_global_styles()

    # ── Sidebar navigation ───────────────────────────────────────────────
    def render_sidebar() -> str:
        with st.sidebar:
            st.markdown(
                """
                <div class="sidebar-brand">
                    <div class="brand-title">SBC BUILDER</div>
                    <div class="brand-subtitle">Data &amp; Reporting Platform</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.divider()

            # Database and Validator have been removed from navigation
            nav_items = {
                "📂 Data Upload":   "Data Upload",
                "🔗 Data Merger":   "Data Merger",
                "📊 Analytics":     "Analytics",
                "🗂️ Template Map":  "Template Map",
                "📄 Reports":       "Reports",
            }

            if "active_page" not in st.session_state or st.session_state.active_page not in nav_items.values():
                st.session_state.active_page = "Data Upload"

            for label, page in nav_items.items():
                active = st.session_state.active_page == page

                if active:
                    st.markdown('<div class="nav-btn-active">', unsafe_allow_html=True)

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

            db_ok = st.session_state.get("db_connected", False)
            status_class = "badge-active" if db_ok else "badge-inactive"
            status_text = "DATABASE CONNECTED" if db_ok else "NO DATABASE"

            st.markdown(
                f"""
                <div style="margin-top:12px;">
                    <span class="badge {status_class}">{status_text}</span>
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

    # ── Page dispatcher ──────────────────────────────────────────────────
    page = render_sidebar()

    # Database and Validator pages removed from dispatch
    dispatch = {
        "Data Upload":  upload_render,
        "Data Merger":  merger_render,
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