"""
modules/db_connector.py — Database connection manager
Supports: PostgreSQL, MySQL, SQLite, MSSQL
"""

import streamlit as st
import pandas as pd
from config.settings import page_header, metric_card, badge

# ── Optional driver imports (graceful degradation) ────────────────────────────
try:
    import sqlalchemy
    from sqlalchemy import create_engine, text, inspect
    _SA_OK = True
except ImportError:
    _SA_OK = False

try:
    import sqlite3 as _sqlite3
    _SQLITE_OK = True
except ImportError:
    _SQLITE_OK = False


# ── Session-state helpers ─────────────────────────────────────────────────────

def _get_engine():
    return st.session_state.get("db_engine", None)


def _get_conn_meta():
    return st.session_state.get("db_meta", {})


def _clear_connection():
    for k in ["db_engine", "db_meta", "db_connected", "db_query_history"]:
        st.session_state.pop(k, None)


# ── Connection builder ────────────────────────────────────────────────────────

def _build_dsn(driver: str, cfg: dict) -> str:
    if driver == "SQLite":
        return f"sqlite:///{cfg['db']}"
    elif driver == "PostgreSQL":
        return (
            f"postgresql+psycopg2://{cfg['user']}:{cfg['password']}"
            f"@{cfg['host']}:{cfg['port']}/{cfg['db']}"
        )
    elif driver == "MySQL":
        return (
            f"mysql+pymysql://{cfg['user']}:{cfg['password']}"
            f"@{cfg['host']}:{cfg['port']}/{cfg['db']}"
        )
    elif driver == "MSSQL":
        return (
            f"mssql+pyodbc://{cfg['user']}:{cfg['password']}"
            f"@{cfg['host']}:{cfg['port']}/{cfg['db']}"
            f"?driver=ODBC+Driver+17+for+SQL+Server"
        )
    raise ValueError(f"Unknown driver: {driver}")


def _try_connect(dsn: str, driver: str, cfg: dict) -> tuple[bool, str]:
    if not _SA_OK:
        return False, "sqlalchemy is not installed."
    try:
        engine = create_engine(dsn, pool_pre_ping=True, connect_args={"connect_timeout": 8})
        with engine.connect() as conn:
            conn.execute(text("SELECT 1"))
        st.session_state["db_engine"] = engine
        st.session_state["db_connected"] = True
        st.session_state["db_meta"] = {
            "driver": driver,
            "host": cfg.get("host", "local"),
            "db": cfg.get("db", ""),
        }
        st.session_state.setdefault("db_query_history", [])
        return True, "Connection successful."
    except Exception as exc:
        return False, str(exc)


# ── Table browser ─────────────────────────────────────────────────────────────

def _render_browser():
    engine = _get_engine()
    if engine is None:
        return
    try:
        insp = inspect(engine)
        tables = insp.get_table_names()
    except Exception as e:
        st.error(f"Could not list tables: {e}")
        return

    if not tables:
        st.info("No tables found in this database.")
        return

    col1, col2 = st.columns([1, 3])
    with col1:
        chosen = st.selectbox("Table", tables)
    with col2:
        row_limit = st.number_input("Row limit", min_value=10, max_value=5000, value=100, step=50)

    if st.button("Preview Table", type="primary"):
        with st.spinner("Fetching…"):
            try:
                with engine.connect() as conn:
                    df = pd.read_sql(
                        text(f'SELECT * FROM "{chosen}" LIMIT :n'),
                        conn, params={"n": int(row_limit)}
                    )
                st.session_state["db_preview"] = df
                st.session_state["db_preview_name"] = chosen
            except Exception as e:
                st.error(str(e))

    if "db_preview" in st.session_state:
        df = st.session_state["db_preview"]
        name = st.session_state.get("db_preview_name", "")
        c1, c2, c3 = st.columns(3)
        c1.metric("Rows", f"{len(df):,}")
        c2.metric("Columns", len(df.columns))
        c3.metric("Table", name)
        st.dataframe(df, use_container_width=True)


# ── Query runner ──────────────────────────────────────────────────────────────

def _render_query():
    engine = _get_engine()
    if engine is None:
        return

    query = st.text_area(
        "SQL Query",
        height=140,
        placeholder="SELECT * FROM my_table WHERE condition LIMIT 100;",
        key="sql_input",
    )

    c1, c2 = st.columns([1, 5])
    run = c1.button("Run Query", type="primary")
    clear_hist = c2.button("Clear History")

    if clear_hist:
        st.session_state["db_query_history"] = []

    if run and query.strip():
        with st.spinner("Executing…"):
            try:
                with engine.connect() as conn:
                    df = pd.read_sql(text(query), conn)
                st.session_state.setdefault("db_query_history", []).append(query.strip())
                st.session_state["db_query_result"] = df
                st.success(f"Returned {len(df):,} rows, {len(df.columns)} columns.")
            except Exception as e:
                st.error(str(e))
                st.session_state.pop("db_query_result", None)

    if "db_query_result" in st.session_state:
        df = st.session_state["db_query_result"]
        st.dataframe(df, use_container_width=True)
        csv = df.to_csv(index=False).encode()
        st.download_button("⬇ Download CSV", csv, "query_result.csv", "text/csv")

    hist = st.session_state.get("db_query_history", [])
    if hist:
        with st.expander(f"Query history ({len(hist)})"):
            for i, q in enumerate(reversed(hist[-20:]), 1):
                st.code(q, language="sql")


# ── Connection panel ──────────────────────────────────────────────────────────

def _render_connect_panel():
    driver = st.selectbox(
        "Database engine",
        ["PostgreSQL", "MySQL", "SQLite", "MSSQL"],
        key="db_driver_sel",
    )

    cfg: dict = {}
    if driver == "SQLite":
        cfg["db"] = st.text_input("File path", value="data.db", placeholder="/path/to/file.db")
    else:
        c1, c2 = st.columns([3, 1])
        cfg["host"] = c1.text_input("Host", placeholder="localhost")
        cfg["port"] = c2.text_input(
            "Port",
            value={"PostgreSQL": "5432", "MySQL": "3306", "MSSQL": "1433"}.get(driver, ""),
        )
        c3, c4 = st.columns(2)
        cfg["user"]     = c3.text_input("Username")
        cfg["password"] = c4.text_input("Password", type="password")
        cfg["db"]       = st.text_input("Database name")

    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
    c_conn, c_disc = st.columns([1, 1])

    with c_conn:
        if st.button("Connect", type="primary", use_container_width=True):
            with st.spinner("Connecting…"):
                dsn = _build_dsn(driver, cfg)
                ok, msg = _try_connect(dsn, driver, cfg)
                if ok:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)

    with c_disc:
        if st.button("Disconnect", use_container_width=True):
            _clear_connection()
            st.rerun()


# ── Main render ───────────────────────────────────────────────────────────────

def render():
    page_header(
        "Database Connection",
        "Connect to a database, browse tables, and run ad-hoc queries.",
    )

    tab_conn, tab_browse, tab_query = st.tabs(["Connection", "Table Browser", "Query Runner"])

    with tab_conn:
        meta = _get_conn_meta()
        if st.session_state.get("db_connected"):
            st.markdown(
                f"""
                <div class="card">
                    <div class="card-title">Active connection</div>
                    <div class="card-value" style="font-size:1.1rem;">
                        {meta.get('driver','–')} · {meta.get('host','–')} · <em>{meta.get('db','–')}</em>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if st.button("Disconnect", type="primary"):
                _clear_connection()
                st.rerun()
        else:
            _render_connect_panel()

    with tab_browse:
        if not st.session_state.get("db_connected"):
            st.info("Connect to a database first.")
        else:
            _render_browser()

    with tab_query:
        if not st.session_state.get("db_connected"):
            st.info("Connect to a database first.")
        else:
            _render_query()
