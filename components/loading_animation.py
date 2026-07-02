"""
Loading Animation Component for SPM Builder
"""
import streamlit as st
import math
from streamlit.components.v1 import html

def create_loading_animation(duration=5):
    """
    Create a loading animation with 5 stars forming a circle around "SBC BUILDER"

    Args:
        duration (int): Duration of the loading animation in seconds (default: 5)

    Returns:
        str: HTML/CSS/JS code for the loading animation
    """

    # Calculate star positions in a circle.
    # Widened radius/container slightly so the longer "SBC BUILDER"
    # text has room and doesn't collide with the stars.
    num_stars = 5
    center_x = 120
    center_y = 120
    radius = 85

    stars_html = ""
    for i in range(num_stars):
        angle = (i / num_stars) * 2 * math.pi - math.pi / 2
        x = center_x + radius * math.cos(angle)
        y = center_y + radius * math.sin(angle)
        delay = i * 0.2
        stars_html += f"""
            <div class="star" style="top: {y-21}px; left: {x-21}px; animation-delay: {delay}s;" id="star{i}">
                ★
            </div>
        """

    loading_html = f"""
    <div id="loadingScreen" class="loading-container">
        <div class="stars-container stars-rotate">
            {stars_html}
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

    <style>
        .loading-container {{
            position: fixed;
            top: 0;
            left: 0;
            width: 100vw;
            height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
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
            color: #FFD700;
            text-shadow: 0 0 20px rgba(255, 215, 0, 0.3);
            transition: all 0.3s ease;
        }}

        .center-text {{
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            text-align: center;
            font-family: 'Arial Black', sans-serif;
            z-index: 10;
            width: 130px;
        }}

        .center-text-line1 {{
            font-size: 15px;
            font-weight: bold;
            color: #FFD700;
            letter-spacing: 5px;
            text-shadow: 0 0 20px rgba(255, 215, 0, 0.3);
            margin-bottom: 2px;
        }}

        .center-text-line2 {{
            font-size: 24px;
            font-weight: bold;
            color: white;
            letter-spacing: 1px;
            text-shadow: 0 0 30px rgba(255, 215, 0, 0.2);
        }}

        .sparkle {{
            animation: sparkle 0.5s ease-in-out infinite alternate;
        }}

        .loading-text {{
            color: #94a3b8;
            font-size: 16px;
            margin-top: 20px;
            font-family: 'Arial', sans-serif;
            letter-spacing: 2px;
            font-weight: 300;
        }}

        @keyframes sparkle {{
            0% {{
                text-shadow: 0 0 10px rgba(255, 215, 0, 0.3);
                transform: scale(1);
            }}
            100% {{
                text-shadow: 0 0 40px rgba(255, 215, 0, 0.9), 0 0 80px rgba(255, 215, 0, 0.5);
                transform: scale(1.2);
            }}
        }}

        @keyframes rotate {{
            0% {{
                transform: rotate(0deg);
            }}
            100% {{
                transform: rotate(360deg);
            }}
        }}

        .stars-rotate {{
            animation: rotate 20s linear infinite;
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
            background: linear-gradient(90deg, #FFD700, #FFA500);
            border-radius: 2px;
            transition: width 0.5s ease;
            width: 0%;
        }}

        .fade-out {{
            animation: fadeOut 0.6s ease forwards;
        }}

        @keyframes fadeOut {{
            0% {{
                opacity: 1;
            }}
            100% {{
                opacity: 0;
                visibility: hidden;
            }}
        }}
    </style>

    <script>
        // Get all stars
        const stars = document.querySelectorAll('.star');
        let progress = 0;
        const progressFill = document.getElementById('progressFill');
        const loadingScreen = document.getElementById('loadingScreen');

        // Function to add sparkle class to random stars
        function sparkleStars() {{
            // Remove all sparkle classes
            stars.forEach(star => star.classList.remove('sparkle'));

            // Add sparkle to random stars
            const numToSparkle = Math.floor(Math.random() * 3) + 1;
            const shuffled = Array.from(stars).sort(() => 0.5 - Math.random());
            for (let i = 0; i < numToSparkle && i < shuffled.length; i++) {{
                shuffled[i].classList.add('sparkle');
            }}
        }}

        // Update progress and trigger sparkles
        function updateProgress() {{
            if (progress < 100) {{
                progress += Math.random() * 3 + 1;
                if (progress > 100) progress = 100;
                progressFill.style.width = progress + '%';
                sparkleStars();

                if (progress < 100) {{
                    setTimeout(updateProgress, 100);
                }} else {{
                    // Complete loading
                    setTimeout(() => {{
                        loadingScreen.classList.add('fade-out');
                        setTimeout(() => {{
                            loadingScreen.style.display = 'none';
                            // Show the main app elements
                            document.querySelector('.stAppHeader').style.display = 'block';
                            document.querySelector('.stSidebar').style.display = 'block';
                            document.querySelector('.stToolbar').style.display = 'block';
                            const mainMenu = document.getElementById('MainMenu');
                            if (mainMenu) mainMenu.style.display = 'block';
                            const footer = document.querySelector('footer');
                            if (footer) footer.style.display = 'block';
                            // Force a rerun to refresh the app
                            window.location.reload();
                        }}, 600);
                    }}, 500);
                }}
            }}
        }}

        // Start the loading animation
        setTimeout(updateProgress, 500);

        // Ensure it ends after max duration
        const maxDuration = {duration * 1000};
        setTimeout(() => {{
            if (progress < 100) {{
                progress = 100;
                progressFill.style.width = '100%';
                loadingScreen.classList.add('fade-out');
                setTimeout(() => {{
                    loadingScreen.style.display = 'none';
                    document.querySelector('.stAppHeader').style.display = 'block';
                    document.querySelector('.stSidebar').style.display = 'block';
                    document.querySelector('.stToolbar').style.display = 'block';
                    const mainMenu = document.getElementById('MainMenu');
                    if (mainMenu) mainMenu.style.display = 'block';
                    const footer = document.querySelector('footer');
                    if (footer) footer.style.display = 'block';
                    window.location.reload();
                }}, 600);
            }}
        }}, maxDuration);
    </script>
    """

    return loading_html


def render_loading_screen(duration=5):
    """
    Render the loading screen in Streamlit

    Args:
        duration (int): Duration of the loading animation in seconds (default: 5)

    Returns:
        bool: True if loading is complete, False otherwise
    """
    if "loading_complete" not in st.session_state:
        st.session_state.loading_complete = False

    if not st.session_state.loading_complete:
        # Hide default Streamlit elements
        st.markdown("""
            <style>
                .stAppHeader, .stSidebar, .stToolbar, #MainMenu, footer {
                    display: none !important;
                }
            </style>
        """, unsafe_allow_html=True)

        # Show loading animation
        st.markdown(create_loading_animation(duration), unsafe_allow_html=True)

        # Mark as complete - will be reset on reload
        st.session_state.loading_complete = True
        return False

    # Show app elements
    st.markdown("""
        <style>
            .stAppHeader { display: block !important; }
            .stSidebar { display: block !important; }
            .stToolbar { display: block !important; }
            #MainMenu { display: block !important; }
            footer { display: block !important; }
        </style>
    """, unsafe_allow_html=True)

    return True


def reset_loading_state():
    """Reset the loading state to show loading screen again"""
    if "loading_complete" in st.session_state:
        st.session_state.loading_complete = False