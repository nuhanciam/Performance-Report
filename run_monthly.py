"""
run_monthly.py — Script CLI pour l'automatisation mensuelle via GitHub Actions.

Variables d'environnement requises :
  SHOPIFY_DOMAIN       ex: nuhanciam.myshopify.com
  SHOPIFY_TOKEN        token admin API Shopify
  GA4_PROPERTY_ID      (optionnel)
  GA4_SERVICE_JSON     (optionnel) — contenu JSON brut du service account GA4

  MAIL_FROM            adresse expéditrice (Gmail recommandé)
  MAIL_PASSWORD        mot de passe application Gmail
  MAIL_TO              destinataire(s) séparés par des virgules

  REPORT_MONTHS        (optionnel) nombre de mois glissants à couvrir, défaut = 1
                       Mettre 12 pour avoir toute l'année dans le même fichier.
"""

import json
import os
import smtplib
import sys
from datetime import date, timedelta
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from analytics_core import (
    build_excel,
    compute_monthly_metrics,
    get_all_orders,
    get_ga4_monthly_metrics,
)


def resolve_period():
    """Retourne (date_from, date_to) pour le mois précédent (comportement par défaut)."""
    nb_months = int(os.environ.get("REPORT_MONTHS", "1"))
    today = date.today()

    # Premier jour du mois courant
    first_current = today.replace(day=1)

    # Dernier jour du mois précédent
    date_to = first_current - timedelta(days=1)

    # Premier jour de la période souhaitée
    month = date_to.month - (nb_months - 1)
    year = date_to.year
    while month <= 0:
        month += 12
        year -= 1
    date_from = date(year, month, 1)

    return date_from, date_to


def send_email(excel_bytes: bytes, filename: str, date_from: date, date_to: date):
    mail_from = os.environ["MAIL_FROM"]
    mail_password = os.environ["MAIL_PASSWORD"]
    mail_to_raw = os.environ["MAIL_TO"]
    recipients = [r.strip() for r in mail_to_raw.split(",") if r.strip()]

    msg = MIMEMultipart()
    msg["From"] = mail_from
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = f"📊 Rapport Shopify Nuhanciam — {date_from} → {date_to}"

    body = (
        f"Bonjour,\n\n"
        f"Veuillez trouver ci-joint le rapport analytique Shopify + GA4 "
        f"pour la période du {date_from} au {date_to}.\n\n"
        f"Ce rapport a été généré automatiquement.\n\n"
        f"Bonne lecture !"
    )
    msg.attach(MIMEText(body, "plain", "utf-8"))

    attachment = MIMEApplication(excel_bytes, _subtype="xlsx")
    attachment.add_header("Content-Disposition", "attachment", filename=filename)
    msg.attach(attachment)

    smtp_host = os.environ.get("SMTP_HOST", "smtp.gmail.com")
    smtp_port = int(os.environ.get("SMTP_PORT", "587"))

    print(f"Envoi du mail à {recipients} via {smtp_host}:{smtp_port}...")

    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.ehlo()
        server.starttls()
        server.login(mail_from, mail_password)
        server.sendmail(mail_from, recipients, msg.as_bytes())

    print("Mail envoyé avec succès ✅")


def main():
    domain = os.environ.get("SHOPIFY_DOMAIN", "nuhanciam.myshopify.com")
    token = os.environ.get("SHOPIFY_TOKEN", "")
    ga4_property_id = os.environ.get("GA4_PROPERTY_ID", "")
    ga4_service_json_raw = os.environ.get("GA4_SERVICE_JSON", "")

    if not token:
        print("❌ SHOPIFY_TOKEN manquant.", file=sys.stderr)
        sys.exit(1)

    if not os.environ.get("MAIL_FROM") or not os.environ.get("MAIL_PASSWORD") or not os.environ.get("MAIL_TO"):
        print("❌ Variables MAIL_FROM / MAIL_PASSWORD / MAIL_TO manquantes.", file=sys.stderr)
        sys.exit(1)

    date_from, date_to = resolve_period()
    print(f"📅 Période : {date_from} → {date_to}")

    # --- Shopify ---
    print("🛍️  Récupération des commandes Shopify...")
    orders = get_all_orders(domain, token, date_from, date_to, log=print)

    if not orders:
        print("Aucune commande trouvée sur cette période. Mail non envoyé.")
        return

    print(f"✅ {len(orders)} commandes récupérées")

    # --- GA4 ---
    ga4_df_result = None
    if ga4_property_id and ga4_service_json_raw:
        print("📈 Récupération des données GA4...")
        try:
            ga4_service_info = json.loads(ga4_service_json_raw)
            import pandas as pd
            ga4_df_result = get_ga4_monthly_metrics(
                property_id=ga4_property_id,
                service_account_info_dict=ga4_service_info,
                date_from=date_from,
                date_to=date_to,
                log=print,
            )
            if ga4_df_result is not None and not ga4_df_result.empty:
                print("✅ Données GA4 récupérées")
            else:
                print("⚠️  Aucune donnée GA4 trouvée")
        except Exception as e:
            print(f"⚠️  Erreur GA4 (rapport généré sans) : {e}")
    else:
        print("ℹ️  GA4 non configuré, lignes analytics vides")

    # --- Calcul ---
    print("🔢 Calcul des métriques...")
    df = compute_monthly_metrics(orders, ga4_df=ga4_df_result, date_from=date_from, date_to=date_to)

    if df.empty:
        print("Aucune commande valide après filtrage. Mail non envoyé.")
        return

    # --- Excel ---
    print("📄 Génération du fichier Excel...")
    excel_bytes = build_excel(df)
    filename = f"analytics_nuhanciam_{date_from}_{date_to}.xlsx"

    # --- Mail ---
    send_email(excel_bytes, filename, date_from, date_to)


if __name__ == "__main__":
    main()
