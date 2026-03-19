# ATS Recap Report Builder

## Overview
Streamlit web app that automates ATS (Available to Ship) recap report generation for Haddad Brands Value Channel Sales team. Users upload raw ATS Excel files pasted from the internal system, configure categories, and download buyer-ready formatted Excel reports.

## Tech Stack
- **Frontend:** Streamlit (matches existing apps: Confirmed Deals Recap, GM Sheet Builder, Store Recap Builder)
- **Excel processing:** openpyxl
- **Image handling:** Pillow (PIL)
- **Auth:** Simple password via `st.secrets["APP_PASSWORD"]`
- **Deployment:** Streamlit Cloud or local

## Project Structure
```
ATS RECAP REPORT/
├── app.py                    # Main Streamlit app (login + upload + configure + generate)
├── requirements.txt
├── CLAUDE.md
├── .streamlit/
│   ├── config.toml           # Theme (cyan/navy matching other Haddad apps)
│   └── secrets.toml          # APP_PASSWORD (not committed)
├── utils/
│   ├── __init__.py
│   ├── auth.py               # Login/logout/require_auth
│   ├── ats_parser.py         # Parse raw ATS Excel, extract images, filter, group
│   └── excel_generator.py    # Generate formatted output Excel with RECAP tab
```

## Business Rules

### Style Number Format
- Example: `76F610-C5E-P3`
- First 6 chars = base style (`76F610`)
- Last 4 of those 6 = REF# (`F610`)
- First digit = size range code:
  - 0=NB Girl, 1=Inf Girl, 2=Tod Girl, 3=4-6X Girl, 4=7-16 Girl
  - 5=NB Boy, 6=Inf Boy, 7=Toddler Boy, 8=4-7 Boy, 9=8-20 Boy

### Filtering
- Remove style/color packs where OH + WIP < 120 units (configurable)
- Optional max threshold (e.g., 12,000 units)
- Re-total at block and category level after filtering

### Raw ATS Structure (per ref# block)
1. Grey header row: STYLE | COLOR | SIZE SCALE | ON HAND | WIP | AVAILABILITY | MSRP
2. Data row pairs (label row with sizes + ratio row with pack quantities)
3. Grey TOTAL row
4. Empty rows / product image area

### Column Layout (Detail Sheets)
- A: Category name (yellow fill, bold) / image area
- B: Spacer / image area
- C: STYLE number
- D: COLOR name
- E-G: Toddler sizes (2T, 3T, 4T)
- H-K: Boys 4-7 sizes (4, 5, 6, 7)
- L: ON HAND
- M: WIP
- N: AVAILABILITY / TOTAL
- O: MSRP

### Output Format
- **Detail sheets:** Yellow category headers, grey sub-headers, TODDLER/4-7 summary rows per category, TOTAL rows per block, embedded product images
- **RECAP tab:** Brand × Category × Size Range summary with:
  - Yellow title + header row
  - Data rows with merged brand/category cells
  - Light blue brand total rows (SUM formulas)
  - Light blue size range total rows (SUMIF formulas)
  - Yellow GRAND TOTAL row

### Image Handling
- Product images (>100px): keep and re-embed
- Color swatches (<100px): remove
- "SWATCH COMING SOON" placeholders: remove

### Categories
- Same category name with different ref#s → merge under one entry
- Omit size range row when no styles exist for that range
- Category order preserved as user defines them

## Running Locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Secrets Required
- `APP_PASSWORD` — shared team password in `.streamlit/secrets.toml`
