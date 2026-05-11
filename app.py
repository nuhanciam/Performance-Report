import json
from datetime import date

import streamlit as st

from analytics_core import (
    build_excel,
    compute_monthly_metrics,
    get_all_orders,
    get_ga4_monthly_metrics,
    make_months_as_columns,
)


st.set_page_config(page_title="Shopify Analytics - Nuhanciam", layout="wide")

st.title("📊 Shopify Analytics - Nuhanciam")
st.markdown("Dashboard mensuel Shopify + GA4 — mois en colonnes, indicateurs en lignes")

with st.sidebar:
    st.header("🔑 Shopify")
    domain_input = st.text_input("Domaine Boutique", value="nuhanciam.myshopify.com")
    token_input = st.text_input("Admin API Token Shopify", type="password")

    st.divider()

    st.header("📈 GA4")
    ga4_property_id = st.text_input("GA4 Property ID")
    ga4_service_file = st.file_uploader("Clé JSON service account GA4", type=["json"])

    st.divider()

    st.subheader("📅 Période")
    col1, col2 = st.columns(2)
    with col1:
        date_from = st.date_input("Du", value=date(date.today().year, 1, 1))
    with col2:
        date_to = st.date_input("Au", value=date.today())

if not token_input:
    st.warning("Veuillez saisir votre Token API Shopify dans la barre latérale.")
    st.stop()

if date_from > date_to:
    st.error("La date de début doit être avant la date de fin.")
    st.stop()

if st.button("🚀 Générer le rapport", use_container_width=True):
    with st.spinner("Récupération des commandes Shopify..."):
        orders = get_all_orders(domain_input, token_input, date_from, date_to)

    if orders is None:
        st.stop()

    if not orders:
        st.warning("Aucune commande trouvée sur cette période.")
        st.stop()

    st.success(f"✅ {len(orders)} commandes Shopify récupérées")

    ga4_df = None
    with st.spinner("Récupération des données GA4..."):
        if ga4_service_file:
            try:
                ga4_service_info = json.load(ga4_service_file)
                ga4_df = get_ga4_monthly_metrics(
                    property_id=ga4_property_id,
                    service_account_info_dict=ga4_service_info,
                    date_from=date_from,
                    date_to=date_to,
                )
            except Exception as e:
                st.warning(f"Erreur GA4 : {e}")

    if ga4_df is None or ga4_df.empty:
        st.info("GA4 non connecté ou aucune donnée trouvée. Les lignes analytics resteront vides.")
    else:
        st.success("✅ Données GA4 récupérées")

    with st.spinner("Calcul des métriques mensuelles..."):
        df = compute_monthly_metrics(orders, ga4_df=ga4_df, date_from=date_from, date_to=date_to)

    if df.empty:
        st.warning("Aucune commande valide trouvée sur cette période après filtrage.")
        st.stop()

    display_df = make_months_as_columns(df)
    st.subheader("📋 Aperçu mensuel")
    st.dataframe(display_df, use_container_width=True, hide_index=True)

    with st.spinner("Génération du fichier Excel..."):
        excel_data = build_excel(df)

    filename = f"analytics_nuhanciam_{date_from}_{date_to}_shopify_ga4.xlsx"
    st.download_button(
        "📥 Télécharger Excel",
        excel_data,
        filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
