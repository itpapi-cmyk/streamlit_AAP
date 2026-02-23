import pandas as pd
from datetime import datetime

from modules.lead_numeric.db import get_conn
from modules.lead_numeric.ddl import init_db

# Excel → DB columns mapping
COLUMN_MAP = {
    "Gruppo": "gruppo",
    "GroupLead": "group_lead",
    "Lead": "lead",
    "Sublead": "sublead",
    "DescrizioneCEE": "descrizione_cee",
    "Tipo": "tipo",
    "SegnoRpt": "segno_rpt",
}

ALLOWED_TIPO = {"ATTIVO", "PASSIVO", "CE"}


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    # strip column names
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in COLUMN_MAP.keys() if c not in df.columns]
    if missing:
        raise ValueError(f"Colonne mancanti nel file schema: {missing}")

    df = df.rename(columns=COLUMN_MAP)
    df = df[list(COLUMN_MAP.values())].copy()

    # normalize text
    df["sublead"] = df["sublead"].astype(str).str.strip()
    df["lead"] = df["lead"].astype(str).str.strip()
    df["group_lead"] = df["group_lead"].astype(str).str.strip()
    df["descrizione_cee"] = df["descrizione_cee"].astype(str).str.strip()
    df["tipo"] = df["tipo"].astype(str).str.strip().str.upper()

    # normalize numbers
    df["segno_rpt"] = pd.to_numeric(df["segno_rpt"], errors="coerce")
    df["gruppo"] = pd.to_numeric(df["gruppo"], errors="coerce")

    return df


def _validate_schema(df: pd.DataFrame) -> None:
    # sublead must exist and be unique
    if df["sublead"].isna().any() or (df["sublead"].str.len() == 0).any():
        raise ValueError("Sublead vuote o non valide nello schema.")

    if df["sublead"].duplicated().any():
        duplicates = df.loc[df["sublead"].duplicated(), "sublead"].unique().tolist()
        raise ValueError(f"Sublead duplicate nello schema: {duplicates[:30]}")

    # tipo allowed
    if not df["tipo"].isin(ALLOWED_TIPO).all():
        invalid = sorted(df.loc[~df["tipo"].isin(ALLOWED_TIPO), "tipo"].unique().tolist())
        raise ValueError(f"Valori 'Tipo' non validi: {invalid}. Ammessi: {sorted(ALLOWED_TIPO)}")

    # segno_rpt numeric and +-1
    if df["segno_rpt"].isna().any():
        raise ValueError("SegnoRpt contiene valori non numerici.")

    unique_signs = set(df["segno_rpt"].unique().tolist())
    if not unique_signs.issubset({1, -1}):
        raise ValueError("SegnoRpt deve contenere solo 1 o -1.")


def import_schema_from_excel(
    excel_path,
    sheet_name=0,
    schema_name="Bilancio UE",
    version="1.0",
    note="Import iniziale schema bilancio",
) -> int:
    """
    Importa lo schema bilancio da Excel nel DB (una tantum per schema_name+version).
    Ritorna schema_version_id.
    """
    init_db()

    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    df = _normalize_columns(df)
    _validate_schema(df)

    conn = get_conn()
    cur = conn.cursor()

    # prevent accidental re-import
    cur.execute(
        "SELECT COUNT(*) FROM lead_schema_version WHERE schema_name=? AND version=?",
        (schema_name, version),
    )
    if cur.fetchone()[0] > 0:
        conn.close()
        raise RuntimeError(f"Schema '{schema_name}' versione {version} già importato.")

    import_date = datetime.now().isoformat(timespec="seconds")

    cur.execute(
        "INSERT INTO lead_schema_version (schema_name, version, import_date, source_file, note) "
        "VALUES (?, ?, ?, ?, ?)",
        (schema_name, version, import_date, str(excel_path), note),
    )
    schema_version_id = cur.lastrowid

    df["schema_version_id"] = schema_version_id
    df.to_sql("lead_structure", conn, if_exists="append", index=False)

    conn.commit()
    conn.close()

    return schema_version_id
