import streamlit as st

def initialize_session():

    defaults = {
        "selected_dataset": None,
        "selected_template": None,
        "logged_user": "Administrator"
    }

    for key, value in defaults.items():

        if key not in st.session_state:
            st.session_state[key] = value