# Shopify Analytics — Nuhanciam

Dashboard Streamlit + automatisation mensuelle GitHub Actions.

## Structure

```
├── analytics_core.py          # Logique métier (Shopify + GA4 + Excel)
├── app.py                     # Interface Streamlit (utilise analytics_core)
├── run_monthly.py             # Script CLI pour l'automatisation
├── .github/
│   └── workflows/
│       └── monthly_report.yml # Cron GitHub Actions (le 1er de chaque mois)
├── data/
│   ├── rapport_cumul.json      # Source de l'Excel cumulatif
│   └── rapport_comparaison.json # Source de l'Excel comparaison mois par mois
└── README.md
```

---

## 1. Mise en place du dépôt GitHub

1. Crée un dépôt privé GitHub (ex: `nuhanciam-analytics`)
2. Pousse tous ces fichiers dedans :
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git remote add origin https://github.com/TON_USER/nuhanciam-analytics.git
   git push -u origin main
   ```

---

## 2. Configurer les Secrets GitHub

Va dans ton dépôt → **Settings → Secrets and variables → Actions → New repository secret**

| Nom du secret        | Valeur                                          |
|----------------------|-------------------------------------------------|
| `SHOPIFY_DOMAIN`     | `nuhanciam.myshopify.com`                       |
| `SHOPIFY_TOKEN`      | Ton token Admin API Shopify                     |
| `MAIL_FROM`          | Ton adresse Gmail (ex: `rapport@gmail.com`)     |
| `MAIL_PASSWORD`      | **Mot de passe d'application** Gmail (voir §3)  |
| `MAIL_TO`            | Destinataire(s) séparés par virgule             |
| `GA4_PROPERTY_ID`    | (optionnel) ID de ta propriété GA4              |
| `GA4_SERVICE_JSON`   | (optionnel) Contenu JSON du service account GA4 |

---

## 3. Créer un mot de passe d'application Gmail

> ⚠️ Ne jamais mettre ton mot de passe Gmail principal — utilise un **App Password**.

1. Active la validation en deux étapes sur ton compte Google
2. Va sur : https://myaccount.google.com/apppasswords
3. Choisis "Mail" → "Autre (nom personnalisé)" → tape "Nuhanciam Analytics"
4. Copie le mot de passe de 16 caractères généré → colle-le dans le secret `MAIL_PASSWORD`

---

## 4. Comportement du workflow

- **Automatique** : se lance le **1er de chaque mois à 7h00 (heure Paris)** et envoie le rapport du mois précédent.
- Le workflow maintient deux JSON : `data/rapport_cumul.json` pour l'Excel cumulatif et `data/rapport_comparaison.json` pour l'Excel de comparaison. Le fichier comparaison est recalculé avec deux requêtes dédiées : le dernier mois complet et le même mois un an avant.
- **Manuel** : tu peux le déclencher à tout moment depuis l'onglet **Actions** de GitHub → bouton "Run workflow". Tu peux choisir combien de mois couvrir (ex: 12 pour toute l'année).

---

## 5. GA4 — Configuration du service account (optionnel)

1. Dans Google Cloud Console, crée un **service account** avec accès en lecture à ta propriété GA4
2. Génère une clé JSON pour ce service account
3. Dans GA4, ajoute l'email du service account comme **Lecteur**
4. Copie le **contenu complet** du fichier JSON et colle-le dans le secret `GA4_SERVICE_JSON`

---

## 6. Lancer Streamlit localement

```bash
pip install streamlit requests pandas openpyxl pytz google-analytics-data google-auth
streamlit run app.py
```
