import streamlit as st
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="SBC Builder - Intelligent Data Processing",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS - Following wiz.py theme
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&family=Inter:wght@300;400;500;600;700&display=swap');

    .stApp {
        background: radial-gradient(circle at 50% -20%, #1c233a 0%, #0d111d 45%, #06080e 100%) !important;
        color: #e2e8f0 !important;
        font-family: 'Inter', sans-serif !important;
    }

    [data-testid="stSidebar"] {
        background-color: #0c0f17 !important;
        border-right: 1px solid rgba(255, 255, 255, 0.05) !important;
    }

    [data-testid="stSidebar"] * {
        font-family: 'Outfit', sans-serif !important;
    }

    .landing-hero {
        font-family: 'Outfit', sans-serif !important;
        background: linear-gradient(135deg, rgba(0, 242, 254, 0.1) 0%, rgba(155, 81, 224, 0.1) 100%) !important;
        backdrop-filter: blur(10px) !important;
        border: 1px solid rgba(0, 242, 254, 0.2) !important;
        border-radius: 16px !important;
        padding: 60px 40px !important;
        margin-bottom: 40px !important;
        text-align: center !important;
        min-height: 400px !important;
        display: flex !important;
        flex-direction: column !important;
        justify-content: center !important;
        align-items: center !important;
    }

    .main-title {
        font-family: 'Outfit', sans-serif !important;
        font-size: 3.5rem !important;
        font-weight: 800 !important;
        background: linear-gradient(135deg, #00f2fe 0%, #4facfe 50%, #9b51e0 100%) !important;
        -webkit-background-clip: text !important;
        -webkit-text-fill-color: transparent !important;
        background-clip: text !important;
        margin-bottom: 0.5rem !important;
        letter-spacing: -1px !important;
    }

    .main-subtitle {
        font-family: 'Outfit', sans-serif !important;
        font-size: 1.3rem !important;
        color: #8a99ad !important;
        margin-bottom: 2rem !important;
        font-weight: 300 !important;
    }

    .cta-button {
        display: inline-block !important;
        background: linear-gradient(135deg, #00f2fe 0%, #4facfe 100%) !important;
        color: #06080e !important;
        padding: 15px 40px !important;
        border-radius: 8px !important;
        font-family: 'Outfit', sans-serif !important;
        font-weight: 600 !important;
        font-size: 1.1rem !important;
        border: none !important;
        cursor: pointer !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 8px 24px 0 rgba(0, 242, 254, 0.3) !important;
    }

    .cta-button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 12px 32px 0 rgba(0, 242, 254, 0.5) !important;
    }

    .feature-card {
        background: linear-gradient(135deg, rgba(255, 255, 255, 0.03) 0%, rgba(255, 255, 255, 0.005) 100%) !important;
        backdrop-filter: blur(16px) !important;
        -webkit-backdrop-filter: blur(16px) !important;
        border: 1px solid rgba(255, 255, 255, 0.05) !important;
        border-top: 1px solid rgba(0, 242, 254, 0.15) !important;
        border-left: 1px solid rgba(0, 242, 254, 0.08) !important;
        border-radius: 12px !important;
        padding: 30px !important;
        margin-bottom: 20px !important;
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.4) !important;
    }

    .feature-icon {
        font-size: 2.5rem !important;
        margin-bottom: 1rem !important;
    }

    .feature-title {
        font-family: 'Outfit', sans-serif !important;
        font-size: 1.5rem !important;
        font-weight: 600 !important;
        color: #00f2fe !important;
        margin-bottom: 0.5rem !important;
    }

    .feature-desc {
        color: #8a99ad !important;
        font-size: 1rem !important;
        line-height: 1.6 !important;
    }

    .section-header {
        font-family: 'Outfit', sans-serif !important;
        font-size: 2rem !important;
        font-weight: 700 !important;
        background: linear-gradient(135deg, #00f2fe 0%, #4facfe 50%, #9b51e0 100%) !important;
        -webkit-background-clip: text !important;
        -webkit-text-fill-color: transparent !important;
        background-clip: text !important;
        margin-bottom: 2rem !important;
        text-align: center !important;
    }

    .glass-divider {
        height: 1px !important;
        background: linear-gradient(to right, transparent, rgba(0, 242, 254, 0.3), transparent) !important;
        margin: 3rem 0 !important;
    }

    .footer-text {
        text-align: center !important;
        color: #4a5568 !important;
        font-size: 0.9rem !important;
        margin-top: 2rem !important;
    }

    </style>
""", unsafe_allow_html=True)

# Hero Section
st.markdown("""
    <div class="landing-hero">
        <div class="main-title">SBC Builder</div>
        <div class="main-subtitle">Intelligent Multi-Sheet Data Processing & Transformation</div>
        <p style="color: #8a99ad; font-size: 1.05rem; margin-bottom: 2.5rem; max-width: 600px;">
            Streamline your data workflows with advanced Excel processing, multi-sheet management, and intelligent data transformation capabilities.
        </p>
    </div>
""", unsafe_allow_html=True)

# CTA Button
col1, col2, col3 = st.columns([1, 0.6, 1])
with col2:
    if st.button("🚀 Get Started", use_container_width=True, key="cta_button"):
        st.session_state.page = "wizard"
        st.rerun()

st.markdown('<div class="glass-divider"></div>', unsafe_allow_html=True)

# Features Section
st.markdown('<div class="section-header">Key Features</div>', unsafe_allow_html=True)

col1, col2, col3 = st.columns(3, gap="medium")

with col1:
    st.markdown("""
        <div class="feature-card">
            <div class="feature-icon">📊</div>
            <div class="feature-title">Multi-Sheet Processing</div>
            <div class="feature-desc">Efficiently handle multiple Excel sheets with intelligent batch operations and automated workflows.</div>
        </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
        <div class="feature-card">
            <div class="feature-icon">🔐</div>
            <div class="feature-title">Secure & Encrypted</div>
            <div class="feature-desc">Full support for encrypted Excel files with password protection and secure data handling.</div>
        </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown("""
        <div class="feature-card">
            <div class="feature-icon">⚡</div>
            <div class="feature-title">Lightning Fast</div>
            <div class="feature-desc">Optimized performance for large datasets with real-time processing and instant feedback.</div>
        </div>
    """, unsafe_allow_html=True)

st.markdown("")

col1, col2, col3 = st.columns(3, gap="medium")

with col1:
    st.markdown("""
        <div class="feature-card">
            <div class="feature-icon">🎨</div>
            <div class="feature-title">Style & Formatting</div>
            <div class="feature-desc">Advanced styling options with cell formatting, colors, and professional design templates.</div>
        </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
        <div class="feature-card">
            <div class="feature-icon">📈</div>
            <div class="feature-title">Data Transformation</div>
            <div class="feature-desc">Powerful data transformation tools including date operations, filtering, and aggregations.</div>
        </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown("""
        <div class="feature-card">
            <div class="feature-icon">📥</div>
            <div class="feature-title">Export & Download</div>
            <div class="feature-desc">Download processed files with format preservation and instant delivery capabilities.</div>
        </div>
    """, unsafe_allow_html=True)

st.markdown('<div class="glass-divider"></div>', unsafe_allow_html=True)

# How It Works Section
st.markdown('<div class="section-header">How It Works</div>', unsafe_allow_html=True)

col1, col2 = st.columns([0.3, 0.7])

with col1:
    st.markdown("""
        <div style="font-family: 'Outfit', sans-serif;">
        <div style="background: linear-gradient(135deg, rgba(0, 242, 254, 0.2) 0%, rgba(79, 172, 254, 0.2) 100%); 
                    border-radius: 50%; 
                    width: 60px; 
                    height: 60px; 
                    display: flex; 
                    align-items: center; 
                    justify-content: center; 
                    font-size: 2rem;
                    margin: 1rem 0;
                    border: 1px solid rgba(0, 242, 254, 0.3);">
            📁
        </div>
        <div style="font-weight: 600; color: #00f2fe; margin-bottom: 0.5rem;">Step 1: Upload</div>
        <div style="color: #8a99ad; font-size: 0.95rem;">Upload your Excel files or spreadsheets</div>
        </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
        <div style="color: #8a99ad; font-size: 1.1rem; padding-top: 2rem;">
        Upload your Excel files with support for multiple formats, encrypted files, and various sheet structures.
        </div>
    """, unsafe_allow_html=True)

st.markdown("")

col1, col2 = st.columns([0.3, 0.7])

with col1:
    st.markdown("""
        <div style="font-family: 'Outfit', sans-serif;">
        <div style="background: linear-gradient(135deg, rgba(79, 172, 254, 0.2) 0%, rgba(155, 81, 224, 0.2) 100%); 
                    border-radius: 50%; 
                    width: 60px; 
                    height: 60px; 
                    display: flex; 
                    align-items: center; 
                    justify-content: center; 
                    font-size: 2rem;
                    margin: 1rem 0;
                    border: 1px solid rgba(79, 172, 254, 0.3);">
            ⚙️
        </div>
        <div style="font-weight: 600; color: #4facfe; margin-bottom: 0.5rem;">Step 2: Configure</div>
        <div style="color: #8a99ad; font-size: 0.95rem;">Set processing parameters and options</div>
        </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
        <div style="color: #8a99ad; font-size: 1.1rem; padding-top: 2rem;">
        Configure your processing workflow with advanced options for data validation, transformation, and output formatting.
        </div>
    """, unsafe_allow_html=True)

st.markdown("")

col1, col2 = st.columns([0.3, 0.7])

with col1:
    st.markdown("""
        <div style="font-family: 'Outfit', sans-serif;">
        <div style="background: linear-gradient(135deg, rgba(155, 81, 224, 0.2) 0%, rgba(0, 242, 254, 0.2) 100%); 
                    border-radius: 50%; 
                    width: 60px; 
                    height: 60px; 
                    display: flex; 
                    align-items: center; 
                    justify-content: center; 
                    font-size: 2rem;
                    margin: 1rem 0;
                    border: 1px solid rgba(155, 81, 224, 0.3);">
            📥
        </div>
        <div style="font-weight: 600; color: #9b51e0; margin-bottom: 0.5rem;">Step 3: Download</div>
        <div style="color: #8a99ad; font-size: 0.95rem;">Get your processed files instantly</div>
        </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
        <div style="color: #8a99ad; font-size: 1.1rem; padding-top: 2rem;">
        Download your processed files with all formatting preserved and ready for immediate use.
        </div>
    """, unsafe_allow_html=True)

st.markdown('<div class="glass-divider"></div>', unsafe_allow_html=True)

# Call to Action
st.markdown("""
    <div style="text-align: center; padding: 2rem;">
        <div style="font-family: 'Outfit', sans-serif; font-size: 1.5rem; font-weight: 600; margin-bottom: 1rem; color: #e2e8f0;">
            Ready to Transform Your Data?
        </div>
        <div style="color: #8a99ad; font-size: 1rem; margin-bottom: 2rem;">
            Start processing your Excel files with SBC Builder today.
        </div>
    </div>
""", unsafe_allow_html=True)

# Footer
st.markdown("""
    <div class="footer-text">
        <p>SBC Builder © 2026 | Intelligent Data Processing Solution</p>
        <p style="font-size: 0.85rem; margin-top: 0.5rem;">Powered by Streamlit | Built with ❤️ for Data Excellence</p>
    </div>
""", unsafe_allow_html=True)
