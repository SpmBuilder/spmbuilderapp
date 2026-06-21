import streamlit as st

def render_metrics():

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric(
            "Data Sources",
            "5",
            "+1"
        )

    with col2:
        st.metric(
            "Datasets",
            "32",
            "+4"
        )

    with col3:
        st.metric(
            "Reports Generated",
            "124",
            "+15"
        )

    with col4:
        st.metric(
            "Failed Jobs",
            "2",
            "-1"
        )