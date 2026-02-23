from modules.lead_numeric.db import get_conn

DDL = """
-- =========================
-- MASTER DATA: SCHEMA LEAD
-- =========================
CREATE TABLE IF NOT EXISTS lead_schema_version (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    schema_name TEXT NOT NULL,
    version TEXT NOT NULL,
    import_date TEXT NOT NULL,
    source_file TEXT NOT NULL,
    note TEXT
);

CREATE TABLE IF NOT EXISTS lead_structure (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    gruppo INTEGER,
    group_lead TEXT,
    lead TEXT,
    sublead TEXT NOT NULL,
    descrizione_cee TEXT,
    tipo TEXT,
    segno_rpt INTEGER,
    schema_version_id INTEGER NOT NULL,
    FOREIGN KEY (schema_version_id) REFERENCES lead_schema_version(id),
    UNIQUE (schema_version_id, sublead)
);

-- =========================
-- CHART OF ACCOUNTS
-- =========================
CREATE TABLE IF NOT EXISTS gl_account (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    account_code TEXT NOT NULL,
    account_name TEXT,
    chart_of_accounts TEXT NOT NULL,
    UNIQUE (account_code, chart_of_accounts)
);

-- =========================
-- MAPPING: CONTO -> SUBLEAD
-- =========================
CREATE TABLE IF NOT EXISTS account_lead_mapping (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    gl_account_id INTEGER NOT NULL,
    sublead TEXT NOT NULL,
    schema_version_id INTEGER NOT NULL,
    is_active INTEGER DEFAULT 1,
    note TEXT,
    FOREIGN KEY (gl_account_id) REFERENCES gl_account(id),
    FOREIGN KEY (schema_version_id) REFERENCES lead_schema_version(id),
    UNIQUE (gl_account_id, schema_version_id, is_active)
);

-- =========================
-- ENTITA' / SOCIETA'
-- =========================
CREATE TABLE IF NOT EXISTS legal_entity (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_code TEXT NOT NULL UNIQUE,
    entity_name TEXT NOT NULL,
    currency TEXT
);

-- =========================
-- TRIAL BALANCE
-- =========================
CREATE TABLE IF NOT EXISTS trial_balance_header (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    legal_entity_id INTEGER NOT NULL,
    fiscal_year INTEGER NOT NULL,
    chart_of_accounts TEXT NOT NULL,
    currency TEXT NOT NULL,
    import_date TEXT NOT NULL,
    source_file TEXT,
    note TEXT,
    UNIQUE (legal_entity_id, fiscal_year, chart_of_accounts),
    FOREIGN KEY (legal_entity_id) REFERENCES legal_entity(id)
);

CREATE TABLE IF NOT EXISTS trial_balance_line (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    trial_balance_id INTEGER NOT NULL,
    gl_account_id INTEGER NOT NULL,
    opening_balance REAL DEFAULT 0,
    debit REAL DEFAULT 0,
    credit REAL DEFAULT 0,
    closing_balance REAL NOT NULL,
    FOREIGN KEY (trial_balance_id) REFERENCES trial_balance_header(id),
    FOREIGN KEY (gl_account_id) REFERENCES gl_account(id),
    UNIQUE (trial_balance_id, gl_account_id)
);
"""

def _migrate_lead_structure_unique_constraint(conn):
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(lead_structure)")
    cols = {row[1]: row for row in cur.fetchall()}
    if not cols:
        return

    # Legacy schema had global UNIQUE on sublead; recreate table with
    # UNIQUE(schema_version_id, sublead) so each version can reuse sublead codes.
    cur.execute("PRAGMA index_list(lead_structure)")
    index_rows = cur.fetchall()
    has_target_unique = False
    has_legacy_sublead_unique = False

    for row in index_rows:
        index_name = row[1]
        is_unique = row[2] == 1
        if not is_unique:
            continue
        cur.execute(f"PRAGMA index_info({index_name!r})")
        index_cols = [i[2] for i in cur.fetchall()]
        if index_cols == ["schema_version_id", "sublead"]:
            has_target_unique = True
        if index_cols == ["sublead"]:
            has_legacy_sublead_unique = True

    if has_target_unique and not has_legacy_sublead_unique:
        return

    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS lead_structure_new (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            gruppo INTEGER,
            group_lead TEXT,
            lead TEXT,
            sublead TEXT NOT NULL,
            descrizione_cee TEXT,
            tipo TEXT,
            segno_rpt INTEGER,
            schema_version_id INTEGER NOT NULL,
            FOREIGN KEY (schema_version_id) REFERENCES lead_schema_version(id),
            UNIQUE (schema_version_id, sublead)
        );
        INSERT INTO lead_structure_new (
            id, gruppo, group_lead, lead, sublead, descrizione_cee, tipo, segno_rpt, schema_version_id
        )
        SELECT
            id, gruppo, group_lead, lead, sublead, descrizione_cee, tipo, segno_rpt, schema_version_id
        FROM lead_structure;
        DROP TABLE lead_structure;
        ALTER TABLE lead_structure_new RENAME TO lead_structure;
        """
    )

def init_db():
    conn = get_conn()
    conn.executescript(DDL)
    _migrate_lead_structure_unique_constraint(conn)
    conn.commit()
    conn.close()
