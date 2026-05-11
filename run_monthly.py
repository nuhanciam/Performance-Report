"""
run_monthly.py — Script CLI pour l'automatisation mensuelle via GitHub Actions.

Fonctionnement :
  - Récupère les données du mois précédent (ou N mois si REPORT_MONTHS > 1)
  - Charge l'Excel cumulatif existant dans le dépôt (data/rapport_cumul.json)
  - Fusionne les nouvelles données avec l'historique
  - Régénère l'Excel complet (toutes les colonnes depuis le début)
  - Sauvegarde le JSON mis à jour dans le dépôt (commité par le workflow)
  - Envoie le mail avec l'Excel en pièce jointe

Variables d'environnement requises :
  SHOPIFY_DOMAIN       ex: nuhanciam.myshopify.com
  SHOPIFY_TOKEN        token admin API Shopify
  GA4_PROPERTY_ID      (optionnel)
  GA4_SERVICE_JSON     (optionnel) — contenu JSON brut du service account GA4

  MAIL_FROM            adresse expéditrice (Gmail recommandé)
  MAIL_PASSWORD        mot de passe application Gmail
  MAIL_TO              destinataire(s) séparés par des virgules

  REPORT_MONTHS        (optionnel) nombre de mois à récupérer ce run, défaut = 1
"""

import json
import os
import smtplib
import sys
from datetime import date, timedelta
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import pandas as pd

from analytics_core import (
    build_excel,
    compute_monthly_metrics,
    get_all_orders,
    get_ga4_monthly_metrics,
)

# Fichier JSON qui stocke l'historique dans le dépôt
CUMUL_PATH = Path("data/rapport_cumul.json")


def resolve_period():
    nb_months = int(os.environ.get("REPORT_MONTHS", "1"))
    today = date.today()
    first_current = today.replace(day=1)
    date_to = first_current - timedelta(days=1)
    month = date_to.month - (nb_months - 1)
    year = date_to.year
    while month <= 0:
        month += 12
        year -= 1
    date_from = date(year, month, 1)
    return date_from, date_to


def load_cumul() -> pd.DataFrame:
    """Charge l'historique cumulatif depuis le JSON du dépôt."""
    if not CUMUL_PATH.exists():
        print("ℹ️  Pas d'historique existant — premier run.")
        return pd.DataFrame()
    try:
        with open(CUMUL_PATH, "r", encoding="utf-8") as f:
            records = json.load(f)
        df = pd.DataFrame(records)
        print(f"📂 Historique chargé : {len(df)} mois existants ({', '.join(df['Mois'].tolist())})")
        return df
    except Exception as e:
        print(f"⚠️  Erreur lecture historique : {e} — repart de zéro.")
        return pd.DataFrame()


def save_cumul(df: pd.DataFrame):
    """Sauvegarde le DataFrame cumulatif en JSON dans le dépôt."""
    CUMUL_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(CUMUL_PATH, "w", encoding="utf-8") as f:
        json.dump(df.to_dict(orient="records"), f, ensure_ascii=False, default=str)
    print(f"💾 Historique sauvegardé : {len(df)} mois au total")


def merge_cumul(existing: pd.DataFrame, new: pd.DataFrame) -> pd.DataFrame:
    """
    Fusionne l'historique existant avec les nouvelles données.
    Si un mois existe déjà, il est remplacé (utile pour corriger un run).
    Les mois sont toujours triés chronologiquement.
    """
    if existing.empty:
        return new.sort_values("Mois").reset_index(drop=True)

    # Supprime les mois déjà présents qui seraient dans 'new' (re-run ou correction)
    new_months = set(new["Mois"].tolist())
    existing_filtered = existing[~existing["Mois"].isin(new_months)]

    merged = pd.concat([existing_filtered, new], ignore_index=True)
    merged = merged.sort_values("Mois").reset_index(drop=True)

    existing_months = set(existing["Mois"].tolist())
    added = new_months - existing_months
    updated = new_months & existing_months

    if added:
        print(f"➕ Nouveaux mois ajoutés : {', '.join(sorted(added))}")
    if updated:
        print(f"🔄 Mois mis à jour : {', '.join(sorted(updated))}")

    return merged


def send_email(excel_bytes: bytes, filename: str, new_month: str, total_months: int):
    mail_from = os.environ["MAIL_FROM"]
    mail_password = os.environ["MAIL_PASSWORD"]
    recipients = [r.strip() for r in os.environ["MAIL_TO"].split(",") if r.strip()]

    msg = MIMEMultipart()
    msg["From"] = mail_from
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = f"📊 Rapport Shopify Nuhanciam — {new_month} ajouté ({total_months} mois au total)"

    body = (
        f"Bonjour,\n\n"
        f"Le rapport mensuel a été mis à jour avec les données de {new_month}.\n"
        f"Le fichier joint contient désormais {total_months} mois d'historique.\n\n"
        f"Ce rapport est généré automatiquement chaque 1er du mois.\n\n"
        f"Bonne lecture !"
    )
    msg.attach(MIMEText(body, "plain", "utf-8"))

    attachment = MIMEApplication(excel_bytes, _subtype="xlsx")
    attachment.add_header("Content-Disposition", "attachment", filename=filename)
    msg.attach(attachment)

    smtp_host = os.environ.get("SMTP_HOST", "smtp.gmail.com")
    smtp_port = int(os.environ.get("SMTP_PORT", "587"))

    print(f"📧 Envoi du mail à {recipients}...")
    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.ehlo()
        server.starttls()
        server.login(mail_from, mail_password)
        server.sendmail(mail_from, recipients, msg.as_bytes())
    print("✅ Mail envoyé avec succès")


def main():
    domain = os.environ.get("SHOPIFY_DOMAIN", "nuhanciam.myshopify.com")
    token = os.environ.get("SHOPIFY_TOKEN", "")
    ga4_property_id = os.environ.get("GA4_PROPERTY_ID", "")
    ga4_service_json_raw = os.environ.get("GA4_SERVICE_JSON", "")

    if not token:
        print("❌ SHOPIFY_TOKEN manquant.", file=sys.stderr)
        sys.exit(1)

    if not all(os.environ.get(v) for v in ["MAIL_FROM", "MAIL_PASSWORD", "MAIL_TO"]):
        print("❌ Variables MAIL_FROM / MAIL_PASSWORD / MAIL_TO manquantes.", file=sys.stderr)
        sys.exit(1)

    date_from, date_to = resolve_period()
    print(f"📅 Période ce run : {date_from} → {date_to}")

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
        print("ℹ️  GA4 non configuré")

    # --- Calcul nouvelles données ---
    print("🔢 Calcul des métriques du mois...")
    new_df = compute_monthly_metrics(
        orders, ga4_df=ga4_df_result, date_from=date_from, date_to=date_to
    )

    if new_df.empty:
        print("Aucune commande valide après filtrage. Mail non envoyé.")
        return

    # --- Fusion avec historique ---
    existing_df = load_cumul()
    full_df = merge_cumul(existing_df, new_df)

    # --- Sauvegarde historique ---
    save_cumul(full_df)

    # --- Génération Excel cumulatif ---
    print(f"📄 Génération de l'Excel ({len(full_df)} mois)...")
    excel_bytes = build_excel(full_df)
    filename = "analytics_nuhanciam_cumul.xlsx"

    # --- Mail ---
    new_months_label = ", ".join(sorted(new_df["Mois"].tolist()))
    send_email(excel_bytes, filename, new_months_label, len(full_df))


if __name__ == "__main__":
    main()