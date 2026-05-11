import json
import time
from io import BytesIO
from collections import defaultdict

import pandas as pd
import requests

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


API_VERSION = "2024-04"
SHOP_TIMEZONE = "Europe/Paris"
REQUEST_DELAY_SECONDS = 0.60
MAX_RETRIES = 6


REPORT_ROWS = [
    ("Sessions", "sessions"),
    ("Visteurs uniques", "unique_visitors"),
    ("Sessions / visitor", "sessions_per_visitor"),
    ("Duration (s)", "duration_seconds"),
    ("Duration (min)", "duration_minutes"),
    ("Bounce %", "bounce_pct"),
    ("", None),
    ("Add to cart", "add_to_cart"),
    ("Checkout", "checkout"),
    ("COMMANDES", "# Commandes"),
    ("Conversion %", "conversion_pct"),
    ("Commandes Nx clients", "# Nouveaux clients"),
    ("% cdes Nx clients", "% cdes Nx clients"),
    ("", None),
    ("# Clients uniques", "# Clients"),
    ("# Nx clients", "# Nouveaux clients"),
    ("", None),
    ("# Clients récurrents", "# Clients récurrents"),
    ("", None),
    ("# Produits vendus", "# Produits vendus"),
    ("Ratio cdes/clients", "Ratio cdes/clients"),
    ("Ratio produit / cde", "Ratio produit/cde"),
    ("Ratio produit / client", "Ratio produit/client"),
    ("", None),
    ("Gross sales", "Gross Sales (€)"),
    ("Discounts", "Discounts (€)"),
    ("Discounts %", "Discounts %"),
    ("", None),
    ("Net Sales", "Net Sales (€)"),
    ("Shipping", "Shipping (€)"),
    ("Taxes", "Taxes (€)"),
    ("TOTAL SALES", "Total Sales (€)"),
    ("POIDS %", "POIDS %"),
    ("", None),
    ("AOV HT", "AOV HT (€)"),
    ("AOV TTC (incl ship)", "AOV TTC incl. ship (€)"),
    ("", None),
    ("CA / client HT", "CA/client HT (€)"),
    ("CA / client TTC (incl ship)", "CA/client TTC (€)"),
    ("CA Nx clients HT", "CA Nvx clients HT (€)"),
    ("% CA Nvx clients", "% CA Nvx clients"),
    ("", None),
    ("Frequence achat", "Fréquence achat"),
    ("LTV estimée", "LTV estimée (€)"),
    ("", None),
    ("Retours €", "Retours (€)"),
    ("Retours #", "Retours (#)"),
    ("Retours %", "Retours %"),
    ("", None),
    ("NET SALES", "Net Sales après retours (€)"),
    ("FRANCE", "France (€)"),
    ("", None),
    ("EXPORT", "Export (€)"),
    ("", None),
    ("TOP 3 EXPORT", None),
    ("#1", "Export #1"),
    ("#2", "Export #2"),
    ("#3", "Export #3"),
    ("", None),
    ("TOP 5 PDCT", None),
    ("#1", "Top Pdct #1"),
    ("#2", "Top Pdct #2"),
    ("#3", "Top Pdct #3"),
    ("#4", "Top Pdct #4"),
    ("#5", "Top Pdct #5"),
]


def clean_domain(domain):
    return domain.strip().replace("https://", "").replace("http://", "").split("/")[0]


def normalize_date(value):
    return pd.to_datetime(value).date()


def shopify_datetime_bounds(date_from, date_to):
    start = pd.Timestamp(date_from).tz_localize(SHOP_TIMEZONE)
    end = (
        pd.Timestamp(date_to) + pd.Timedelta(hours=23, minutes=59, seconds=59)
    ).tz_localize(SHOP_TIMEZONE)
    return start.isoformat(), end.isoformat()


def parse_order_created_at(order):
    created_at = order.get("created_at")
    if not created_at:
        return None
    try:
        return pd.to_datetime(created_at, utc=True).tz_convert(SHOP_TIMEZONE)
    except Exception:
        return None


def get_next_link(link_header):
    if not link_header:
        return None
    for part in link_header.split(","):
        chunks = part.split(";")
        if len(chunks) < 2:
            continue
        url_part = chunks[0].strip()
        rel_part = ";".join(chunks[1:])
        if 'rel="next"' in rel_part:
            return url_part.strip("<> ")
    return None


def request_shopify_with_retry(session, url, headers, params=None, log=print):
    last_response = None
    for attempt in range(MAX_RETRIES + 1):
        try:
            response = session.get(url, headers=headers, params=params, timeout=30)
        except requests.RequestException as e:
            if attempt == MAX_RETRIES:
                raise e
            wait = min(2 ** attempt, 10)
            log(f"Connexion instable. Nouvelle tentative dans {wait:.1f}s...")
            time.sleep(wait)
            continue

        last_response = response

        if response.status_code == 429 and attempt < MAX_RETRIES:
            retry_after = response.headers.get("Retry-After")
            try:
                wait = float(retry_after) if retry_after else min(2 ** attempt + 1, 10)
            except ValueError:
                wait = min(2 ** attempt + 1, 10)
            wait = max(wait, REQUEST_DELAY_SECONDS)
            log(f"Rate limit 429. Pause de {wait:.1f}s...")
            time.sleep(wait)
            continue

        if 500 <= response.status_code < 600 and attempt < MAX_RETRIES:
            wait = min(2 ** attempt + 1, 10)
            log(f"Erreur Shopify {response.status_code}. Nouvelle tentative dans {wait:.1f}s...")
            time.sleep(wait)
            continue

        return response

    return last_response


def get_all_orders(domain, token, date_from, date_to, log=print):
    domain = clean_domain(domain)
    created_at_min, created_at_max = shopify_datetime_bounds(date_from, date_to)

    url = f"https://{domain}/admin/api/{API_VERSION}/orders.json"
    headers = {"X-Shopify-Access-Token": token.strip()}
    params = {
        "limit": 250,
        "status": "any",
        "created_at_min": created_at_min,
        "created_at_max": created_at_max,
    }

    orders = []
    session = requests.Session()

    while url:
        try:
            response = request_shopify_with_retry(
                session=session, url=url, headers=headers, params=params, log=log
            )
        except Exception as e:
            raise RuntimeError(f"Erreur connexion Shopify : {e}")

        if response.status_code == 200:
            data = response.json()
            orders.extend(data.get("orders", []))
            log(f"{len(orders)} commandes récupérées...")
            url = get_next_link(response.headers.get("Link", ""))
            params = None
            if url:
                time.sleep(REQUEST_DELAY_SECONDS)
        else:
            raise RuntimeError(f"Erreur API Shopify ({response.status_code}): {response.text[:300]}")

    return orders


def get_ga4_monthly_metrics(property_id, service_account_info_dict, date_from, date_to, log=print):
    if not property_id or not service_account_info_dict:
        return pd.DataFrame()

    try:
        from google.analytics.data_v1beta import BetaAnalyticsDataClient
        from google.analytics.data_v1beta.types import DateRange, Dimension, Metric, RunReportRequest
        from google.oauth2 import service_account
    except ImportError:
        log("Librairies GA4 manquantes (google-analytics-data google-auth)")
        return pd.DataFrame()

    try:
        credentials = service_account.Credentials.from_service_account_info(
            service_account_info_dict,
            scopes=["https://www.googleapis.com/auth/analytics.readonly"],
        )
        client = BetaAnalyticsDataClient(credentials=credentials)
        request = RunReportRequest(
            property=f"properties/{property_id.strip()}",
            dimensions=[Dimension(name="yearMonth")],
            metrics=[
                Metric(name="sessions"),
                Metric(name="totalUsers"),
                Metric(name="sessionsPerUser"),
                Metric(name="averageSessionDuration"),
                Metric(name="bounceRate"),
                Metric(name="addToCarts"),
                Metric(name="checkouts"),
            ],
            date_ranges=[DateRange(start_date=str(date_from), end_date=str(date_to))],
        )
        response = client.run_report(request)
    except Exception as e:
        log(f"GA4 non récupéré : {e}")
        return pd.DataFrame()

    rows = []
    for row in response.rows:
        year_month = row.dimension_values[0].value
        month_key = f"{year_month[:4]}-{year_month[4:]}"
        values = [metric.value for metric in row.metric_values]
        sessions = float(values[0] or 0)
        duration_seconds = float(values[3] or 0)
        bounce_rate = float(values[4] or 0)
        rows.append({
            "Mois": month_key,
            "sessions": round(sessions),
            "unique_visitors": round(float(values[1] or 0)),
            "sessions_per_visitor": round(float(values[2] or 0), 2),
            "duration_seconds": round(duration_seconds, 1),
            "duration_minutes": round(duration_seconds / 60, 2),
            "bounce_pct": round(bounce_rate * 100, 1),
            "add_to_cart": round(float(values[5] or 0)),
            "checkout": round(float(values[6] or 0)),
        })

    return pd.DataFrame(rows)


def compute_monthly_metrics(orders, ga4_df=None, date_from=None, date_to=None):
    date_from = normalize_date(date_from) if date_from else None
    date_to = normalize_date(date_to) if date_to else None

    valid_orders = []
    for order in orders:
        if order.get("financial_status") in ("voided",):
            continue
        created_local = parse_order_created_at(order)
        if created_local is None:
            continue
        order_date = created_local.date()
        if date_from and order_date < date_from:
            continue
        if date_to and order_date > date_to:
            continue
        month_key = created_local.strftime("%Y-%m")
        valid_orders.append((order, month_key))

    monthly = defaultdict(lambda: {"orders": [], "customers": set(), "returning_customers": set()})
    customer_order_count = defaultdict(int)

    for order, _ in valid_orders:
        customer = order.get("customer") or {}
        customer_id = customer.get("id")
        if customer_id:
            customer_order_count[customer_id] += 1

    for order, month_key in valid_orders:
        monthly[month_key]["orders"].append(order)
        customer = order.get("customer") or {}
        customer_id = customer.get("id")
        if customer_id:
            monthly[month_key]["customers"].add(customer_id)
            if customer_order_count[customer_id] > 1:
                monthly[month_key]["returning_customers"].add(customer_id)

    rows = []
    for month_key in sorted(monthly.keys()):
        month_data = monthly[month_key]
        month_orders = month_data["orders"]
        nb_orders = len(month_orders)
        nb_customers = len(month_data["customers"])
        nb_returning = len(month_data["returning_customers"])
        new_customers = nb_customers - nb_returning

        gross_sales = sum(float(o.get("subtotal_price") or 0) for o in month_orders)
        discounts = sum(float(o.get("total_discounts") or 0) for o in month_orders)
        shipping = sum(
            float(o.get("total_shipping_price_set", {}).get("shop_money", {}).get("amount", 0) or 0)
            for o in month_orders
        )
        taxes = sum(float(o.get("total_tax") or 0) for o in month_orders)
        total_sales = sum(float(o.get("total_price") or 0) for o in month_orders)
        net_sales = gross_sales - discounts

        refund_amount = 0
        refund_count = 0
        for order in month_orders:
            for refund in order.get("refunds", []):
                for refund_item in refund.get("refund_line_items", []):
                    refund_amount += float(refund_item.get("subtotal") or 0)
                    refund_count += int(refund_item.get("quantity") or 0)

        products_sold = sum(
            int(line_item.get("quantity") or 0)
            for order in month_orders
            for line_item in order.get("line_items", [])
        )

        france_sales = 0
        export_sales = 0
        country_sales = defaultdict(float)
        for order in month_orders:
            address = order.get("shipping_address") or order.get("billing_address") or {}
            country = address.get("country_code", "").upper()
            amount = float(order.get("total_price") or 0)
            if country == "FR":
                france_sales += amount
            else:
                export_sales += amount
                if country:
                    country_sales[country] += amount

        top3_export = sorted(country_sales.items(), key=lambda x: x[1], reverse=True)[:3]

        product_qty = defaultdict(lambda: {"title": "", "qty": 0, "revenue": 0.0})
        for order in month_orders:
            for line_item in order.get("line_items", []):
                product_id = str(line_item.get("product_id", ""))
                quantity = int(line_item.get("quantity") or 0)
                price = float(line_item.get("price") or 0)
                product_qty[product_id]["title"] = line_item.get("title", product_id)
                product_qty[product_id]["qty"] += quantity
                product_qty[product_id]["revenue"] += price * quantity

        top5 = sorted(product_qty.values(), key=lambda x: x["revenue"], reverse=True)[:5]

        ratio_cde_client = round(nb_orders / nb_customers, 2) if nb_customers else 0
        ratio_pdct_cde = round(products_sold / nb_orders, 2) if nb_orders else 0
        ratio_pdct_client = round(products_sold / nb_customers, 2) if nb_customers else 0
        aov_ht = round(net_sales / nb_orders, 2) if nb_orders else 0
        aov_ttc = round(total_sales / nb_orders, 2) if nb_orders else 0
        ca_client_ht = round(net_sales / nb_customers, 2) if nb_customers else 0
        ca_client_ttc = round(total_sales / nb_customers, 2) if nb_customers else 0
        ca_new_clients = round((new_customers / nb_customers) * net_sales, 2) if nb_customers else 0
        pct_ca_new = round((new_customers / nb_customers) * 100, 1) if nb_customers else 0
        ltv = round(ca_client_ht * ratio_cde_client * 12, 2)
        refund_pct = round((refund_amount / gross_sales) * 100, 1) if gross_sales else 0
        discount_pct = round((discounts / gross_sales) * 100, 1) if gross_sales else 0

        rows.append({
            "Mois": month_key,
            "sessions": "", "unique_visitors": "", "sessions_per_visitor": "",
            "duration_seconds": "", "duration_minutes": "", "bounce_pct": "",
            "add_to_cart": "", "checkout": "", "conversion_pct": "",
            "# Clients": nb_customers,
            "# Clients récurrents": nb_returning,
            "# Nouveaux clients": new_customers,
            "# Commandes": nb_orders,
            "# Produits vendus": products_sold,
            "Ratio cdes/clients": ratio_cde_client,
            "Ratio produit/cde": ratio_pdct_cde,
            "Ratio produit/client": ratio_pdct_client,
            "Gross Sales (€)": round(gross_sales, 2),
            "Discounts (€)": round(discounts, 2),
            "Discounts %": discount_pct,
            "Net Sales (€)": round(net_sales, 2),
            "Shipping (€)": round(shipping, 2),
            "Taxes (€)": round(taxes, 2),
            "Total Sales (€)": round(total_sales, 2),
            "AOV HT (€)": aov_ht,
            "AOV TTC incl. ship (€)": aov_ttc,
            "CA/client HT (€)": ca_client_ht,
            "CA/client TTC (€)": ca_client_ttc,
            "CA Nvx clients HT (€)": ca_new_clients,
            "% CA Nvx clients": pct_ca_new,
            "Fréquence achat": ratio_cde_client,
            "LTV estimée (€)": ltv,
            "Retours (€)": round(refund_amount, 2),
            "Retours (#)": refund_count,
            "Retours %": refund_pct,
            "Net Sales après retours (€)": round(net_sales - refund_amount, 2),
            "France (€)": round(france_sales, 2),
            "Export (€)": round(export_sales, 2),
            "Export #1 pays": top3_export[0][0] if len(top3_export) > 0 else "",
            "Export #1 CA (€)": round(top3_export[0][1], 2) if len(top3_export) > 0 else 0,
            "Export #2 pays": top3_export[1][0] if len(top3_export) > 1 else "",
            "Export #2 CA (€)": round(top3_export[1][1], 2) if len(top3_export) > 1 else 0,
            "Export #3 pays": top3_export[2][0] if len(top3_export) > 2 else "",
            "Export #3 CA (€)": round(top3_export[2][1], 2) if len(top3_export) > 2 else 0,
            "Top Pdct #1": top5[0]["title"] if len(top5) > 0 else "",
            "Top Pdct #1 CA (€)": round(top5[0]["revenue"], 2) if len(top5) > 0 else 0,
            "Top Pdct #2": top5[1]["title"] if len(top5) > 1 else "",
            "Top Pdct #2 CA (€)": round(top5[1]["revenue"], 2) if len(top5) > 1 else 0,
            "Top Pdct #3": top5[2]["title"] if len(top5) > 2 else "",
            "Top Pdct #3 CA (€)": round(top5[2]["revenue"], 2) if len(top5) > 2 else 0,
            "Top Pdct #4": top5[3]["title"] if len(top5) > 3 else "",
            "Top Pdct #4 CA (€)": round(top5[3]["revenue"], 2) if len(top5) > 3 else 0,
            "Top Pdct #5": top5[4]["title"] if len(top5) > 4 else "",
            "Top Pdct #5 CA (€)": round(top5[4]["revenue"], 2) if len(top5) > 4 else 0,
        })

    df = pd.DataFrame(rows)

    if not df.empty:
        total_period_sales = df["Total Sales (€)"].sum()
        df["POIDS %"] = df["Total Sales (€)"].apply(
            lambda v: round((v / total_period_sales) * 100, 1) if total_period_sales else 0
        )
        df["% cdes Nx clients"] = df.apply(
            lambda row: round((row["# Nouveaux clients"] / row["# Commandes"]) * 100, 1)
            if row["# Commandes"] else 0, axis=1,
        )
        for i in range(1, 4):
            df[f"Export #{i}"] = df.apply(
                lambda row, i=i: f'{row[f"Export #{i} pays"]} - {row[f"Export #{i} CA (€)"]} €'
                if row[f"Export #{i} pays"] else "", axis=1,
            )

    if ga4_df is not None and not ga4_df.empty and not df.empty:
        df = df.merge(ga4_df, on="Mois", how="left", suffixes=("", "_ga4"))
        ga4_columns = [
            "sessions", "unique_visitors", "sessions_per_visitor",
            "duration_seconds", "duration_minutes", "bounce_pct",
            "add_to_cart", "checkout",
        ]
        for col in ga4_columns:
            ga4_col = f"{col}_ga4"
            if ga4_col in df.columns:
                df[col] = df[ga4_col].fillna("")
                df = df.drop(columns=[ga4_col])

        df["conversion_pct"] = df.apply(
            lambda row: round((row["# Commandes"] / row["sessions"]) * 100, 2)
            if isinstance(row["sessions"], (int, float)) and row["sessions"] else "",
            axis=1,
        )

    return df


def make_months_as_columns(df):
    if df.empty:
        return pd.DataFrame()
    df_by_month = df.set_index("Mois")
    months = list(df_by_month.index)
    rows = []
    for label, source_key in REPORT_ROWS:
        row = {"Indicateur": label}
        for month in months:
            if source_key and source_key in df_by_month.columns:
                row[month] = df_by_month.at[month, source_key]
            else:
                row[month] = ""
        rows.append(row)
    return pd.DataFrame(rows)


def build_excel(df):
    display_df = make_months_as_columns(df)

    wb = Workbook()
    ws = wb.active
    ws.title = "Vue mensuelle"

    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    normal_font = Font(name="Arial", size=10)
    bold_font = Font(name="Arial", bold=True, size=10)
    header_fill = PatternFill("solid", fgColor="1F4E79")
    section_fill = PatternFill("solid", fgColor="D9E1F2")
    stripe_fill = PatternFill("solid", fgColor="F2F2F2")
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    euro_fmt = '#,##0.00 "€"'
    pct_fmt = '0.0"%"'
    num_fmt = '#,##0'
    ratio_fmt = '0.00'

    euro_labels = {
        "Gross sales", "Discounts", "Net Sales", "Shipping", "Taxes",
        "TOTAL SALES", "AOV HT", "AOV TTC (incl ship)", "CA / client HT",
        "CA / client TTC (incl ship)", "CA Nx clients HT", "LTV estimée",
        "Retours €", "NET SALES", "FRANCE", "EXPORT",
    }
    pct_labels = {
        "Bounce %", "Conversion %", "% cdes Nx clients", "Discounts %",
        "POIDS %", "% CA Nvx clients", "Retours %",
    }
    ratio_labels = {
        "Sessions / visitor", "Ratio cdes/clients", "Ratio produit / cde",
        "Ratio produit / client", "Frequence achat",
    }
    count_labels = {
        "Sessions", "Visteurs uniques", "Add to cart", "Checkout",
        "COMMANDES", "Commandes Nx clients", "# Clients uniques",
        "# Nx clients", "# Clients récurrents", "# Produits vendus", "Retours #",
    }
    section_labels = {"TOP 3 EXPORT", "TOP 5 PDCT"}

    headers = list(display_df.columns)

    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    for row_idx, row in display_df.iterrows():
        excel_row = row_idx + 2
        label = row["Indicateur"]

        for col_idx, header in enumerate(headers, start=1):
            value = row[header]
            cell = ws.cell(row=excel_row, column=col_idx, value=value)

            if label == "":
                cell.border = Border()
                continue

            cell.border = border
            cell.alignment = Alignment(
                horizontal="left" if col_idx == 1 else "right", vertical="center"
            )

            if label in section_labels:
                cell.fill = section_fill
                cell.font = Font(name="Arial", bold=True, color="1F4E79", size=11)
                continue

            cell.font = bold_font if col_idx == 1 else normal_font

            if row_idx % 2 == 0:
                cell.fill = stripe_fill

            if col_idx > 1:
                if label in euro_labels:
                    cell.number_format = euro_fmt
                elif label in pct_labels:
                    cell.number_format = pct_fmt
                elif label in ratio_labels:
                    cell.number_format = ratio_fmt
                elif label in count_labels:
                    cell.number_format = num_fmt

    ws.freeze_panes = "B2"
    ws.column_dimensions["A"].width = 34
    for col_idx in range(2, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 16

    output = BytesIO()
    wb.save(output)
    return output.getvalue()
