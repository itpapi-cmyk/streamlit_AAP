import io

import pandas as pd
import streamlit as st

from modules.lead_numeric.db import get_conn
from modules.lead_numeric.ddl import init_db

st.set_page_config(page_title="Mapping conti -> Sublead", layout="wide")
st.title("03 - Mapping conti -> Sublead")

init_db()


def do_rerun():
    try:
        st.rerun()
    except Exception:
        st.experimental_rerun()


def _latest_schema_id(conn):
    value = pd.read_sql(
        "SELECT MAX(id) AS schema_version_id FROM lead_schema_version",
        conn,
    ).iloc[0]["schema_version_id"]
    if pd.isna(value):
        return None
    return int(value)


def _tb_list(conn):
    return pd.read_sql(
        """
        SELECT tbh.id, le.entity_code, tbh.fiscal_year, tbh.chart_of_accounts, tbh.import_date
        FROM trial_balance_header tbh
        JOIN legal_entity le ON le.id = tbh.legal_entity_id
        ORDER BY tbh.import_date DESC, tbh.id DESC
        """,
        conn,
    )


MAPPING_CTE = """
WITH latest_schema AS (
    SELECT MAX(id) AS id FROM lead_schema_version
),
latest_mapping AS (
    SELECT gl_account_id, sublead
    FROM (
        SELECT m.*,
               ROW_NUMBER() OVER (
                   PARTITION BY m.gl_account_id
                   ORDER BY m.schema_version_id DESC, m.id DESC
               ) AS rn
        FROM account_lead_mapping m
        WHERE m.is_active = 1
    )
    WHERE rn = 1
),
valid_mapping AS (
    SELECT lm.gl_account_id, lm.sublead
    FROM latest_mapping lm
    JOIN latest_schema s ON 1=1
    JOIN lead_structure ls
      ON ls.schema_version_id = s.id
     AND ls.sublead = lm.sublead
)
"""


try:
    conn = get_conn()

    latest_schema_id = _latest_schema_id(conn)
    if latest_schema_id is None:
        st.error("Nessuna versione schema trovata. Importa prima lo schema in pagina 01.")
        st.stop()

    tb_headers = _tb_list(conn)
    if tb_headers.empty:
        st.warning("Nessun Trial Balance importato. Vai in pagina 02 per importare un TB.")
        st.stop()

    tb_options = [
        f"{int(r.id)} | {r.entity_code} | {int(r.fiscal_year)} | {r.chart_of_accounts} | {r.import_date}"
        for _, r in tb_headers.iterrows()
    ]
    selected_tb_opt = st.selectbox(
        "Perimetro mapping (Trial Balance)",
        options=tb_options,
        index=0,
        help="La pagina mostra solo i conti presenti nel Trial Balance selezionato.",
    )
    selected_tb_id = int(selected_tb_opt.split("|")[0].strip())

    df_sublead = pd.read_sql(
        """
        SELECT ls.sublead, ls.lead, ls.group_lead, ls.descrizione_cee, ls.tipo
        FROM lead_structure ls
        WHERE ls.schema_version_id = ?
        ORDER BY ls.tipo, ls.gruppo, ls.sublead
        """,
        conn,
        params=(latest_schema_id,),
    )
    if df_sublead.empty:
        st.error("Non trovo sublead nello schema. Hai importato lo schema bilancio in pagina 01?")
        st.stop()

    sublead_options = [
        f"{r.sublead} | {r.lead} | {r.descrizione_cee}"
        for _, r in df_sublead.iterrows()
    ]
    opt_to_sublead = dict(zip(sublead_options, df_sublead["sublead"].tolist()))

    df_unmapped = pd.read_sql(
        MAPPING_CTE
        + """
        SELECT ga.id AS gl_account_id, ga.account_code, ga.account_name
        FROM trial_balance_line tbl
        JOIN gl_account ga ON ga.id = tbl.gl_account_id
        LEFT JOIN valid_mapping vm ON vm.gl_account_id = ga.id
        WHERE tbl.trial_balance_id = ?
          AND vm.gl_account_id IS NULL
        GROUP BY ga.id, ga.account_code, ga.account_name
        ORDER BY ga.account_code
        """,
        conn,
        params=(selected_tb_id,),
    )

    st.caption("Assegna i conti non mappati a una Sublead. Dopo il salvataggio spariscono dalla lista.")
    if df_unmapped.empty:
        st.success("Nessun conto da mappare nel Trial Balance selezionato.")
    else:
        st.warning(f"Conti da mappare nel TB selezionato: {len(df_unmapped)}")

    with st.expander("Riepilogo conti mappati", expanded=False):
        df_mapped = pd.read_sql(
            MAPPING_CTE
            + """
            SELECT ga.account_code, ga.account_name, vm.sublead, ls.lead, ls.descrizione_cee
            FROM trial_balance_line tbl
            JOIN gl_account ga ON ga.id = tbl.gl_account_id
            JOIN valid_mapping vm ON vm.gl_account_id = ga.id
            JOIN lead_structure ls
              ON ls.schema_version_id = ?
             AND ls.sublead = vm.sublead
            WHERE tbl.trial_balance_id = ?
            GROUP BY ga.id, ga.account_code, ga.account_name, vm.sublead, ls.lead, ls.descrizione_cee
            ORDER BY ga.account_code
            """,
            conn,
            params=(latest_schema_id, selected_tb_id),
        )

        if df_mapped.empty:
            st.info("Nessun conto mappato.")
        else:
            st.dataframe(df_mapped, use_container_width=True)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df_mapped.to_excel(writer, index=False, sheet_name="Mapping")
            st.download_button(
                label="Esporta in Excel",
                data=output.getvalue(),
                file_name="riepilogo_mapping.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    st.subheader("Allinea o modifica conti a Sublead (selezione multipla)")
    col1, col2 = st.columns(2)

    df_all = pd.read_sql(
        MAPPING_CTE
        + """
        SELECT ga.id AS gl_account_id, ga.account_code, ga.account_name, vm.sublead AS mapped_sublead
        FROM trial_balance_line tbl
        JOIN gl_account ga ON ga.id = tbl.gl_account_id
        LEFT JOIN valid_mapping vm ON vm.gl_account_id = ga.id
        WHERE tbl.trial_balance_id = ?
        GROUP BY ga.id, ga.account_code, ga.account_name, vm.sublead
        ORDER BY ga.account_code
        """,
        conn,
        params=(selected_tb_id,),
    )
    df_all["account_name"] = df_all["account_name"].fillna("")
    df_all["label"] = df_all["account_code"] + " - " + df_all["account_name"]

    with col2:
        st.caption("Seleziona la Sublead a cui allineare i conti:")
        selected_sublead_opt = st.selectbox("Sublead", options=sublead_options)
        selected_sublead = opt_to_sublead[selected_sublead_opt]

        df_assigned = df_all[df_all["mapped_sublead"] == selected_sublead].copy()
        if not df_assigned.empty:
            st.write(f"Conti gia assegnati a Sublead {selected_sublead}:")
            df_assigned["label"] = df_assigned["account_code"] + " - " + df_assigned["account_name"]
            assigned_labels = df_assigned["label"].tolist()
            label_to_id = dict(zip(df_assigned["label"], df_assigned["gl_account_id"]))
            remove_accounts = st.multiselect(
                "Seleziona conti da rimuovere da questa Sublead",
                options=assigned_labels,
                default=[],
            )
            if remove_accounts and st.button("Rimuovi conti selezionati da Sublead"):
                error_count = 0
                for acc_label in remove_accounts:
                    try:
                        acc_id = int(label_to_id[acc_label])
                        res = conn.execute(
                            """
                            DELETE FROM account_lead_mapping
                              WHERE gl_account_id = ?
                                AND sublead = ?
                                AND is_active = 1
                            """,
                            (acc_id, selected_sublead),
                        )
                        if res.rowcount == 0:
                            error_count += 1
                    except Exception:
                        error_count += 1
                conn.commit()
                if error_count == 0:
                    st.success("Conti rimossi dalla Sublead.")
                else:
                    st.warning("Alcuni conti non sono stati rimossi correttamente.")
                do_rerun()

    df_unmapped_only = df_all[df_all["mapped_sublead"].isnull()]
    account_options = df_unmapped_only["label"].tolist()
    account_ids = dict(zip(df_unmapped_only["label"], df_unmapped_only["gl_account_id"]))

    with col1:
        st.caption("Seleziona uno o piu conti da mappare:")
        selected_accounts = st.multiselect("Conti", options=account_options, default=[])

    if selected_accounts and st.button("Allinea conti selezionati alla Sublead"):
        for acc_label in selected_accounts:
            acc_id = int(account_ids[acc_label])
            conn.execute(
                """
                INSERT INTO account_lead_mapping
                (gl_account_id, sublead, schema_version_id, is_active)
                VALUES (?, ?, ?, 1)
                """,
                (acc_id, selected_sublead, latest_schema_id),
            )
        conn.commit()
        st.success(f"Conti allineati a Sublead {selected_sublead}")
        do_rerun()

except Exception as e:
    st.error("Errore nella pagina Mapping (03). Dettaglio:")
    st.exception(e)

finally:
    try:
        conn.close()
    except Exception:
        pass
