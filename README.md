# 🎯 Sales Ops · Lead Classifier
**Delivery Hero / Pandora · Digital Sales APAC — Singapore**

Classifies new restaurant leads against Salesforce CRM and Google Maps to identify duplicates, potential matches, and genuinely new business opportunities.

---

## Features

- **CRM deduplication** — matches leads against Salesforce All Accounts by postal code and name similarity
- **Google Maps validation** — uses Apify scrape results to confirm business status and category eligibility
- **Chinese name support** — automatic Hanyu Pinyin conversion for accurate matching of Chinese restaurant names
- **SF Account Audit** — finds suspected duplicate accounts within Salesforce itself
- **Configurable thresholds** — P3/P4 match sensitivity adjustable from the sidebar

---

## Classification Labels

| Label | Meaning | Action |
|---|---|---|
| **P1 — New** | No CRM match. Google confirms open restaurant. | ✅ Pitch delivery |
| **P3 — Potential Match** | Name 50–74% similar to CRM account at same postal. | 🔍 Verify manually |
| **P4 — Duplicate** | Name ≥ 75% similar to CRM account at same postal. | ❌ Skip |
| **Business Closed** | Google Maps shows permanently / temporarily closed. | ❌ Skip |
| **Wrong Target Group** | Google Maps category not food delivery eligible. | ❌ Skip |
| **P2 — Please Check** | No Google Maps result or unreliable match. | ⚪ Manual check |

---

## Requirements

```
streamlit>=1.32.0
pandas>=2.0.0
openpyxl>=3.1.0
xlsxwriter>=3.1.0
shapely>=2.0.0
pypinyin>=0.50.0
rapidfuzz>=3.0.0
```

Install all dependencies:
```bash
pip install -r requirements.txt
```

---

## Secrets

Create `.streamlit/secrets.toml` with:

```toml
APP_PASSWORD   = "your_password"
ONEMAP_EMAIL   = "your_onemap_email"
ONEMAP_PASSWORD = "your_onemap_password"
```

OneMap credentials are used for Singapore postal code geocoding (delivery zone checking). Register at [onemap.gov.sg](https://www.onemap.gov.sg).

---

## Files Required

### 1. Leads File (from Salesforce)
Export from the Salesforce Leads Report. Required columns:
- `GRID`
- `Company / Account`
- `Street`
- `Zip/Postal Code`
- `Coordinates (Latitude)` *(optional — used for zone checking)*
- `Coordinates (Longitude)` *(optional)*

### 2. Apify Results File (from Google Maps Scraper)
Export from Apify after running Google Maps Scraper on the generated URLs. Required columns:
- `GRID` *(add manually — copy from leads file in same row order)*
- `title`
- `categoryName`
- `permanentlyClosed`
- `temporarilyClosed`
- `street`, `phone`, `website`, `url` *(optional enrichment)*

### 3. CRM Export (Salesforce All Accounts)
Export from the Singapore CRM Report. Required columns:
- `GRID`
- `Account Name`
- `Account Status`
- `Restaurant PostalCode`
- `Formatted Restaurant Address`

---

## Workflow

```
1. Export Leads from Salesforce
        ↓
2. Generate Apify URLs  (Tab 2 → Generate Apify URLs)
        ↓
3. Run Apify Google Maps Scraper
   Add GRID column to Apify export
        ↓
4. Export CRM All Accounts from Salesforce
        ↓
5. Upload all 3 files → Run Classification  (Tab 1 → Classify Leads)
        ↓
6. Download Excel report with 6 labelled sheets
```

---

## Matching Logic

Leads are matched against CRM using the following cascade:

1. **Postal code** — only CRM accounts at the same 6-digit SG postal are compared
2. **Unit number** — if the lead has a unit (e.g. `#01-23`), all CRM accounts at that exact unit are scored. All accounts at the same unit are compared (not just the first)
3. **Name similarity** — `rapidfuzz.token_set_ratio` + Hanyu Pinyin comparison for Chinese names
   - Score ≥ 75% → **P4 Duplicate**
   - Score 50–74% → **P3 Potential Match**
   - Score < 50% → no CRM match → Apify check

If no CRM match is found, the lead is validated against Apify:
- Found + food category + open → **P1 New**
- Found + closed + address mismatch → **P2 Please Check**
- Found + closed + address match → **Business Closed**
- Found + non-food category + name mismatch → **P2 Please Check**
- Found + non-food category → **Wrong Target Group**
- Not found or no category → **P2 Please Check**

---

## Tabs

| Tab | Purpose |
|---|---|
| 📊 Classify Leads | Main classification tool |
| 🔗 Generate Apify URLs | Generate Google Maps search URLs for Apify input |
| 🏢 SF Account Audit | Find duplicate accounts within Salesforce itself |
| 📖 How to Use | In-app usage guide |

---

## Excel Report Sheets

| Sheet | Contents |
|---|---|
| Classified Leads | All leads, colour-coded by label |
| Summary | Breakdown by label, match method, top categories |
| ✅ P1 — New | Sales-ready leads with previous occupant context |
| 🔴 P4 — Duplicate | Confirmed duplicates with CRM match details |
| 🟡 P3 — Potential | Leads needing rep verification |
| ⚪ P2 — Please Check | Unverifiable leads |
| ⚠️ Closed + Wrong TG | Business Closed and Wrong Target Group leads |

---

## Optional: Delivery Zone Checking

Place a `zones_SG.json` file in the same directory as `app.py`. Format:
```json
[
  {
    "zone_name": "Zone A",
    "city_name": "Central",
    "wkt": "POLYGON((103.8 1.28, 103.85 1.28, ...))"
  }
]
```

Leads will be checked against these polygons using OneMap geocoding.

---

## Running Locally

```bash
streamlit run app.py
```

---

*Internal tool · Sales Ops · Digital Sales APAC · Delivery Hero / Pandora*
