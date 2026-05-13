"""
run_monthly.py — Script CLI pour l'automatisation mensuelle via GitHub Actions.

Fonctionnement :
  - Récupère les données du mois précédent (ou N mois si REPORT_MONTHS > 1)
  - Charge l'Excel cumulatif existant dans le dépôt (data/rapport_cumul.json)
  - Fusionne les nouvelles données avec l'historique
  - Régénère l'Excel complet (toutes les colonnes depuis le début)
  - Sauvegarde les JSON mis à jour dans le dépôt (commités par le workflow)
  - Envoie le mail avec les deux Excel en pièces jointes

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
import builtins
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
    build_excel_comparison,
    compute_monthly_metrics,
    get_all_orders,
    get_ga4_monthly_metrics,
)


def safe_print(*args, **kwargs):
    """Print safely when a local console cannot encode every character."""
    stream = kwargs.get("file") or sys.stdout
    encoding = getattr(stream, "encoding", None) or "utf-8"
    safe_args = [
        str(arg).encode(encoding, errors="replace").decode(encoding, errors="replace")
        for arg in args
    ]
    builtins.print(*safe_args, **kwargs)


print = safe_print

# Fichiers JSON qui stockent les historiques dans le dépôt
CUMUL_PATH = Path("data/rapport_cumul.json")
COMPARISON_PATH = Path("data/rapport_comparaison.json")


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


def month_bounds(year: int, month: int):
    date_from = date(year, month, 1)
    if month == 12:
        date_to = date(year + 1, 1, 1) - timedelta(days=1)
    else:
        date_to = date(year, month + 1, 1) - timedelta(days=1)
    return date_from, date_to


def resolve_comparison_periods(reference_date_to: date):
    current_from, current_to = month_bounds(reference_date_to.year, reference_date_to.month)
    previous_from, previous_to = month_bounds(current_from.year - 1, current_from.month)
    return (current_from, current_to), (previous_from, previous_to)


def load_history(path: Path, label: str) -> pd.DataFrame:
    """Charge un historique depuis un JSON du dépôt."""
    if not path.exists():
        print(f"ℹ️  Pas d'historique {label} existant ({path}) — premier run.")
        return pd.DataFrame()
    try:
        with open(path, "r", encoding="utf-8") as f:
            records = json.load(f)
        df = pd.DataFrame(records)

        if df.empty:
            print(f"📂 Historique {label} chargé : 0 mois existant")
            return df

        if "Mois" not in df.columns:
            raise ValueError("colonne 'Mois' manquante")

        df = df.sort_values("Mois").reset_index(drop=True)
        print(f"📂 Historique {label} chargé : {len(df)} mois existants ({', '.join(df['Mois'].tolist())})")
        return df
    except Exception as e:
        print(f"⚠️  Erreur lecture historique {label} : {e} — repart de zéro.")
        return pd.DataFrame()


def save_history(path: Path, df: pd.DataFrame, label: str):
    """Sauvegarde un DataFrame historique en JSON dans le dépôt."""
    path.parent.mkdir(parents=True, exist_ok=True)
    df_to_save = df.sort_values("Mois").reset_index(drop=True) if "Mois" in df.columns else df
    with open(path, "w", encoding="utf-8") as f:
        json.dump(df_to_save.to_dict(orient="records"), f, ensure_ascii=False, default=str)
    print(f"💾 Historique {label} sauvegardé : {len(df_to_save)} mois au total ({path})")


def load_cumul() -> pd.DataFrame:
    """Charge l'historique cumulatif depuis le JSON du dépôt."""
    return load_history(CUMUL_PATH, "cumulatif")


def save_cumul(df: pd.DataFrame):
    """Sauvegarde l'historique cumulatif en JSON dans le dépôt."""
    save_history(CUMUL_PATH, df, "cumulatif")


def load_comparison_source() -> pd.DataFrame:
    """Charge l'historique utilisé pour générer l'Excel de comparaison."""
    return load_history(COMPARISON_PATH, "comparaison")


def save_comparison_source(df: pd.DataFrame):
    """Sauvegarde le JSON source dédié à l'Excel de comparaison."""
    save_history(COMPARISON_PATH, df, "comparaison")


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


def fetch_metrics_for_period(
    domain: str,
    token: str,
    ga4_property_id: str,
    ga4_service_info: dict | None,
    date_from: date,
    date_to: date,
    label: str,
) -> pd.DataFrame:
    """Récupère Shopify/GA4 puis calcule les métriques mensuelles pour une période."""
    print(f"🛍️  Récupération des commandes Shopify ({label}) : {date_from} → {date_to}")
    orders = get_all_orders(domain, token, date_from, date_to, log=print)

    if not orders:
        print(f"⚠️  Aucune commande trouvée pour {label}.")
        return pd.DataFrame()

    print(f"✅ {len(orders)} commandes récupérées pour {label}")

    ga4_df_result = None
    if ga4_property_id and ga4_service_info:
        print(f"📈 Récupération des données GA4 ({label})...")
        try:
            ga4_df_result = get_ga4_monthly_metrics(
                property_id=ga4_property_id,
                service_account_info_dict=ga4_service_info,
                date_from=date_from,
                date_to=date_to,
                log=print,
            )
            if ga4_df_result is not None and not ga4_df_result.empty:
                print(f"✅ Données GA4 récupérées pour {label}")
            else:
                print(f"⚠️  Aucune donnée GA4 trouvée pour {label}")
        except Exception as e:
            print(f"⚠️  Erreur GA4 pour {label} (rapport généré sans) : {e}")
    else:
        print(f"ℹ️  GA4 non configuré pour {label}")

    print(f"🔢 Calcul des métriques ({label})...")
    return compute_monthly_metrics(
        orders, ga4_df=ga4_df_result, date_from=date_from, date_to=date_to
    )


def send_email(
    excel_cumul_bytes: bytes,
    filename_cumul: str,
    excel_comp_bytes: bytes,
    filename_comp: str,
    new_month: str,
    total_months: int,
    comparison_month: str | None = None,
):
    mail_from = os.environ["MAIL_FROM"]
    mail_password = os.environ["MAIL_PASSWORD"]
    recipients = [r.strip() for r in os.environ["MAIL_TO"].split(",") if r.strip()]

    msg = MIMEMultipart()
    msg["From"] = mail_from
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = f"📊 Rapport Shopify Nuhanciam — {new_month} ({total_months} mois d'historique)"
    comparison_month = comparison_month or new_month

    body = (
        f"Bonjour,\n\n"
        f"Le rapport mensuel a été mis à jour avec les données de {new_month}.\n\n"
        f"2 fichiers joints :\n"
        f"  • {filename_cumul} — historique complet ({total_months} mois, une colonne par mois)\n"
        f"  • {filename_comp} — comparaison {comparison_month} vs même mois l'année précédente\n\n"
        f"Ce rapport est généré automatiquement chaque 1er du mois.\n\n"
        f"Bonne lecture !"
    )
    msg.attach(MIMEText(body, "plain", "utf-8"))

    for excel_bytes, filename in [
        (excel_cumul_bytes, filename_cumul),
        (excel_comp_bytes, filename_comp),
    ]:
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
    ga4_service_info = None

    if not token:
        print("❌ SHOPIFY_TOKEN manquant.", file=sys.stderr)
        sys.exit(1)

    if not all(os.environ.get(v) for v in ["MAIL_FROM", "MAIL_PASSWORD", "MAIL_TO"]):
        print("❌ Variables MAIL_FROM / MAIL_PASSWORD / MAIL_TO manquantes.", file=sys.stderr)
        sys.exit(1)

    if ga4_property_id and ga4_service_json_raw:
        try:
            ga4_service_info = json.loads(ga4_service_json_raw)
        except Exception as e:
            print(f"⚠️  Erreur lecture GA4_SERVICE_JSON (rapports générés sans GA4) : {e}")

    date_from, date_to = resolve_period()
    print(f"📅 Période ce run : {date_from} → {date_to}")

    new_df = fetch_metrics_for_period(
        domain=domain,
        token=token,
        ga4_property_id=ga4_property_id,
        ga4_service_info=ga4_service_info,
        date_from=date_from,
        date_to=date_to,
        label="historique cumulatif",
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
    print(f"📄 Génération de l'Excel cumulatif ({len(full_df)} mois)...")
    excel_cumul_bytes = build_excel(full_df)
    filename_cumul = "analytics_nuhanciam_cumul.xlsx"

    # --- Génération JSON comparaison depuis deux requêtes dédiées ---
    (comp_current_from, comp_current_to), (comp_previous_from, comp_previous_to) = (
        resolve_comparison_periods(date_to)
    )
    month_m = comp_current_from.strftime("%Y-%m")
    month_m1 = comp_previous_from.strftime("%Y-%m")

    comparison_current_df = fetch_metrics_for_period(
        domain=domain,
        token=token,
        ga4_property_id=ga4_property_id,
        ga4_service_info=ga4_service_info,
        date_from=comp_current_from,
        date_to=comp_current_to,
        label=f"comparaison {month_m}",
    )
    if comparison_current_df.empty and "Mois" in new_df.columns:
        comparison_current_df = new_df[new_df["Mois"] == month_m].copy()

    comparison_previous_df = fetch_metrics_for_period(
        domain=domain,
        token=token,
        ga4_property_id=ga4_property_id,
        ga4_service_info=ga4_service_info,
        date_from=comp_previous_from,
        date_to=comp_previous_to,
        label=f"comparaison {month_m1}",
    )

    comparison_df = pd.concat(
        [comparison_previous_df, comparison_current_df],
        ignore_index=True,
    )
    if not comparison_df.empty and "Mois" in comparison_df.columns:
        comparison_df = (
            comparison_df.drop_duplicates(subset=["Mois"], keep="last")
            .sort_values("Mois")
            .reset_index(drop=True)
        )
    save_comparison_source(comparison_df)

    # --- Génération Excel comparaison depuis son JSON dédié ---
    comparison_df = load_comparison_source()
    if comparison_df.empty:
        print("⚠️  JSON de comparaison vide après sauvegarde — fallback sur les mois disponibles dans le cumulatif.")
        comparison_df = full_df[full_df["Mois"].isin([month_m1, month_m])].copy()

    # --- Génération Excel comparaison (Mois M vs même mois année précédente) ---
    print(f"📊 Génération de l'Excel comparaison pour {month_m}...")
    try:
        excel_comp_bytes = build_excel_comparison(
            comparison_df, month_m, months_to_compare=[month_m]
        )
        filename_comp = f"analytics_nuhanciam_comparaison_{month_m}.xlsx"
    except Exception as e:
        print(f"⚠️  Erreur génération comparaison : {e} — seul le cumulatif sera envoyé.")
        excel_comp_bytes = excel_cumul_bytes
        filename_comp = filename_cumul

    # --- Mail ---
    new_months_label = ", ".join(sorted(new_df["Mois"].tolist()))
    send_email(
        excel_cumul_bytes, filename_cumul,
        excel_comp_bytes, filename_comp,
        new_months_label, len(full_df),
        comparison_month=month_m,
    )


if __name__ == "__main__":
    main()
