import pandas as pd
from datetime import datetime
from modules.lead_numeric.db import get_conn
from modules.lead_numeric.ddl import init_db

# TB columns in your file
REQUIRED_COLS = {"conto", "descrizione", "dare", "avere"}


def import_trial_balance_from_excel(
    excel_file,
    sheet_name=0,
    entity_code="ENTITY01",
    entity_name="ENTITY01",
    fiscal_year=2024,
    chart_of_accounts="COA",
    currency="EUR",
    source_file_name="uploaded.xlsx",
):
    """
    Import TB con colonne: conto, descrizione, dare, avere.
    opening = 0 (se non presente).
    closing = 0 + dare - avere
    Ritorna: (trial_balance_id, unmapped_df)
    """
    init_db()

    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"Colonne mancanti TB: {missing}. Attese: {sorted(REQUIRED_COLS)}")

    # normalize
    df["conto"] = df["conto"].astype(str).str.strip()
    df["descrizione"] = df["descrizione"].astype(str).str.strip()
    df["dare"] = pd.to_numeric(df["dare"], errors="coerce").fillna(0.0)
    df["avere"] = pd.to_numeric(df["avere"], errors="coerce").fillna(0.0)

    df["opening"] = 0.0
    df["closing"] = df["opening"] + df["dare"] - df["avere"]

    conn = get_conn()
    cur = conn.cursor()

    # legal entity
    cur.execute(
        "INSERT OR IGNORE INTO legal_entity (entity_code, entity_name, currency) VALUES (?, ?, ?)",
        (entity_code, entity_name, currency),
    )
    cur.execute("SELECT id FROM legal_entity WHERE entity_code=?", (entity_code,))
    legal_entity_id = cur.fetchone()[0]

    # TB header (unico per entity+year+coa) - upsert
    import_date = datetime.now().isoformat(timespec="seconds")
    cur.execute(
        """
        INSERT OR REPLACE INTO trial_balance_header
        (id, legal_entity_id, fiscal_year, chart_of_accounts, currency, import_date, source_file, note)
        VALUES (
            (SELECT id FROM trial_balance_header WHERE legal_entity_id=? AND fiscal_year=? AND chart_of_accounts=?),
            ?, ?, ?, ?, ?, ?, ?
        )
        """,
        (
            legal_entity_id, fiscal_year, chart_of_accounts,
            legal_entity_id, fiscal_year, chart_of_accounts, currency,
            import_date, source_file_name,
            "Import TB da UI Streamlit",
        ),
    )

    cur.execute(
        "SELECT id FROM trial_balance_header WHERE legal_entity_id=? AND fiscal_year=? AND chart_of_accounts=?",
        (legal_entity_id, fiscal_year, chart_of_accounts),
    )
    trial_balance_id = cur.fetchone()[0]

    # re-import: pulizia righe
    cur.execute("DELETE FROM trial_balance_line WHERE trial_balance_id=?", (trial_balance_id,))

    # insert accounts + tb lines
    for _, r in df.iterrows():
        acc_code = r["conto"]
        acc_name = r["descrizione"]

        cur.execute(
            """
            INSERT OR IGNORE INTO gl_account (account_code, account_name, chart_of_accounts)
            VALUES (?, ?, ?)
            """,
            (acc_code, acc_name, chart_of_accounts),
        )

        cur.execute(
            "SELECT id FROM gl_account WHERE account_code=? AND chart_of_accounts=?",
            (acc_code, chart_of_accounts),
        )
        gl_account_id = cur.fetchone()[0]

        cur.execute(
            """
            INSERT INTO trial_balance_line
            (trial_balance_id, gl_account_id, opening_balance, debit, credit, closing_balance)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                trial_balance_id,
                gl_account_id,
                float(r["opening"]),
                float(r["dare"]),
                float(r["avere"]),
                float(r["closing"]),
            ),
        )

    conn.commit()

    unmapped_df = pd.read_sql(
        """
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
        SELECT ga.id AS gl_account_id, ga.account_code, ga.account_name
        FROM trial_balance_line tbl
        JOIN gl_account ga ON ga.id = tbl.gl_account_id
        LEFT JOIN valid_mapping vm ON vm.gl_account_id = ga.id
        WHERE tbl.trial_balance_id = ?
          AND vm.gl_account_id IS NULL
        ORDER BY ga.account_code
        """,
        conn,
        params=(trial_balance_id,),
    )

    conn.close()
    return trial_balance_id, unmapped_df
