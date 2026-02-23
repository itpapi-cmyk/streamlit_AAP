import streamlit as st
from modules.lead_numeric.ddl import init_db
from modules.lead_numeric.import_tb import import_trial_balance_from_excel

st.title("02 — Import Trial Balance (debit/credit)")
init_db()

col1, col2, col3 = st.columns(3)
with col1:
    entity_code = st.text_input("Entity code", value="ENTITY01")
    entity_name = st.text_input("Entity name", value="ENTITY01")
with col2:
    fiscal_year = st.number_input("Fiscal year", min_value=2000, max_value=2100, value=2024, step=1)
    currency = st.text_input("Currency", value="EUR")
with col3:
    chart_of_accounts = st.text_input("Chart of accounts", value="COA")
    sheet = st.text_input("Nome foglio TB (vuoto = primo)", value="")

uploaded = st.file_uploader("Carica TB Excel", type=["xlsx"])

st.caption("Formato atteso colonne: account_code, account_name, opening, debit, credit")

if uploaded and st.button("📥 Importa TB"):
    sheet_name = 0 if sheet.strip() == "" else sheet.strip()
    try:
        tb_id, unmapped = import_trial_balance_from_excel(
            excel_file=uploaded,
            sheet_name=sheet_name,
            entity_code=entity_code,
            entity_name=entity_name,
            fiscal_year=int(fiscal_year),
            chart_of_accounts=chart_of_accounts,
            currency=currency,
            source_file_name=getattr(uploaded, "name", "uploaded.xlsx"),
        )
        st.success(f"TB importato ✅ (trial_balance_id={tb_id})")

        if len(unmapped) > 0:
            st.warning(f"⚠️ Conti non mappati: {len(unmapped)}. Vai alla pagina 03 — Mapping conti.")
            st.dataframe(unmapped, use_container_width=True)
        else:
            st.success("🎉 Tutti i conti risultano già mappati.")
    except Exception as e:
        st.error(str(e))
