import streamlit as st

def show_dashboard():

    st.title("📊 Dashboard")

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Data Sources", 0)
    col2.metric("Datasets", 0)
    col3.metric("Reports", 0)
    col4.metric("Failed Jobs", 0)

    st.divider()

    left, right = st.columns([2, 1])

    with left:
        st.subheader("Recent Activities")

        st.dataframe(
            [],
            use_container_width=True
        )

    with right:
        st.subheader("System Status")

        st.success("Application Online")
        st.info("Database Not Configured")