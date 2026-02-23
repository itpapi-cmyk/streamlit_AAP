import pandas as pd
import streamlit as st

from modules.lead_numeric.db import get_conn
from modules.lead_numeric.ddl import init_db
from modules.lead_numeric.import_schema import import_schema_from_excel

st.title("01 - Setup Schema Bilancio")

init_db()


def _load_schema_table(schema_version_id=None) -> pd.DataFrame:
    conn = get_conn()
    try:
        if schema_version_id is None:
            latest = pd.read_sql_query(
                """
                SELECT id AS schema_version_id
                FROM lead_schema_version
                ORDER BY id DESC
                LIMIT 1
                """,
                conn,
            )
            if latest.empty:
                return pd.DataFrame()
            schema_version_id = int(latest.iloc[0]["schema_version_id"])

        return pd.read_sql_query(
            """
            SELECT
                s.schema_name,
                s.version,
                l.gruppo,
                l.group_lead,
                l.lead,
                l.sublead,
                l.descrizione_cee,
                l.tipo,
                l.segno_rpt
            FROM lead_structure l
            JOIN lead_schema_version s ON s.id = l.schema_version_id
            WHERE l.schema_version_id = ?
            ORDER BY l.gruppo, l.group_lead, l.lead, l.sublead
            """,
            conn,
            params=(schema_version_id,),
        )
    finally:
        conn.close()


st.subheader("Import una tantum (schema bilancio)")
schema_name = st.text_input("Nome schema", value="Bilancio UE")
version = st.text_input("Versione", value="1.0")
sheet = st.text_input("Nome foglio (opzionale, lascia vuoto = primo foglio)", value="")

uploaded = st.file_uploader("Carica Excel schema bilancio", type=["xlsx"])
imported_schema_version_id = None

if uploaded:
    st.info("File caricato. Premi Import per caricare lo schema nel DB.")
    if st.button("Importa schema nel DB"):
        sheet_name = 0 if sheet.strip() == "" else sheet.strip()
        try:
            schema_version_id = import_schema_from_excel(
                excel_path=uploaded,
                sheet_name=sheet_name,
                schema_name=schema_name,
                version=version,
                note="Import da UI Streamlit",
            )
            imported_schema_version_id = schema_version_id
            st.success(f"Schema importato (schema_version_id={schema_version_id})")
        except Exception as e:
            st.error(str(e))

st.subheader("Tabella schema bilancio")
df_schema = _load_schema_table(imported_schema_version_id)
if df_schema.empty:
    st.info("Nessuno schema presente nel DB.")
else:
    st.dataframe(df_schema, use_container_width=True)
