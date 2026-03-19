"""
ATS Recap Report Builder — Main Application
Haddad Brands Value Channel Sales

Upload raw ATS Excel -> auto-detect categories -> clean -> download formatted report.
"""

import logging
from datetime import date

import streamlit as st

from utils.auth import login, check_password, logout
from utils.ats_parser import parse_ats_file, filter_categories, detect_brand
from utils.excel_generator import generate_ats_report
from utils.security import (
    sanitize_text, sanitize_for_excel, validate_xlsx_file, check_rate_limit,
)

logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(name)s: %(message)s')
logger = logging.getLogger(__name__)

st.set_page_config(page_title="ATS Recap Report Builder", page_icon="📊",
                   layout="wide", initial_sidebar_state="collapsed")

# ─── Static CSS ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
#MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}
.login-card { background: white; border-radius: 16px; padding: 48px 40px 40px;
    box-shadow: 0 8px 32px rgba(27,43,75,0.12); max-width: 420px; margin: 60px auto 0; }
.login-logo { font-size: 2.8rem; text-align: center; margin-bottom: 6px; }
.login-title { font-size: 1.5rem; font-weight: 700; color: #1B2B4B; text-align: center; }
.login-subtitle { font-size: 0.9rem; color: #6B7C99; text-align: center; margin-bottom: 32px; }
.stTextInput > div > div > input { border-radius: 8px !important; border: 1.5px solid #D0D7E5 !important; }
.stTextInput > div > div > input:focus { border-color: #00B0F0 !important; }
.stButton > button { width: 100%; background: linear-gradient(135deg, #00B0F0, #0080C0) !important;
    color: white !important; border: none !important; border-radius: 8px !important;
    padding: 12px !important; font-weight: 600 !important; }
.stButton > button:hover { opacity: 0.9; }
</style>
""", unsafe_allow_html=True)


# ─── Login ───────────────────────────────────────────────────────────────────
if not check_password():
    st.markdown("""<div class="login-card"><div class="login-logo">📊</div>
      <div class="login-title">ATS Recap Report Builder</div>
      <div class="login-subtitle">Value Channel Sales</div></div>""", unsafe_allow_html=True)
    col_l, col_m, col_r = st.columns([1, 2, 1])
    with col_m:
        st.write("")
        name = st.text_input("Your Name", placeholder="e.g. Jacob H.", key="login_name", max_chars=50)
        password = st.text_input("Password", type="password", placeholder="Enter team password",
                                  key="login_pw", max_chars=100)
        if st.button("Sign In", key="login_btn"):
            if not name.strip():
                st.error("Please enter your name.")
            elif not login(name, password):
                st.error("Incorrect password or too many attempts.")
            else:
                st.rerun()
    st.stop()

# ─── Main App ────────────────────────────────────────────────────────────────
user_name = st.session_state.get('user_name', '')
st.title("ATS Recap Report Builder")
st.caption(f"Upload → Configure → Download  |  Welcome, {user_name}")

# ─── Step 1: Upload ──────────────────────────────────────────────────────────
st.subheader("Step 1: Upload Raw ATS File")
uploaded_file = st.file_uploader("Upload your raw ATS Excel file (.xlsx)", type=["xlsx"], key="ats_upload")

if uploaded_file is None:
    st.info("Upload your raw ATS Excel file to get started. Categories are auto-detected from yellow headers.")
    st.stop()

file_bytes = uploaded_file.read()
is_valid, error_msg = validate_xlsx_file(file_bytes, uploaded_file.name)
if not is_valid:
    st.error(f"File validation failed: {error_msg}")
    st.stop()

if not check_rate_limit("upload", max_requests=30, window_seconds=3600):
    st.error("Too many uploads. Please wait.")
    st.stop()

# ─── Parse ────────────────────────────────────────────────────────────────────
# NOTE: Do NOT use @st.cache_data here — it serializes return values,
# which strips binary image bytes from the parsed blocks.
@st.cache_resource(ttl=300, show_spinner="Parsing file...")
def cached_parse(fb):
    return parse_ats_file(fb)

try:
    parsed = cached_parse(file_bytes)
except Exception:
    logger.exception("Parse failed")
    st.error("Could not parse the file. Verify it is a valid ATS Excel file.")
    st.stop()

# Check if any sheets have categories
sheets_with_data = [s for s in parsed["sheets"] if s["categories"]]
if not sheets_with_data:
    st.error("No categories detected. Make sure your raw ATS file has yellow-highlighted category names in column A.")
    st.stop()

# ─── Step 2: Review Auto-Detected Categories ─────────────────────────────────
st.subheader("Step 2: Review Detected Categories")

raw_title = st.text_input("Report Title",
                           value=uploaded_file.name.replace(".xlsx", "").replace(".xls", "").upper(),
                           key="report_title", max_chars=200)
report_title = sanitize_for_excel(sanitize_text(raw_title, max_length=200))
report_date = st.date_input("Report Date", value=date.today(), key="report_date")

st.markdown("---")

sheet_configs = {}

for sheet_idx, sheet in enumerate(sheets_with_data):
    sheet_name = sheet["name"]
    brand = sheet["brand"]
    categories = sheet["categories"]

    cat_names = [c["name"] for c in categories]
    total_blocks = sum(len(c["blocks"]) for c in categories)
    total_refs = len(sheet["all_ref_nums"])

    with st.expander(f"📋 {sheet_name} — {len(cat_names)} categories, {total_refs} ref#s", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            raw_brand = st.text_input("Brand", value=brand or "UNKNOWN",
                                       key=f"brand_{sheet_idx}", max_chars=50)
            sheet_brand = sanitize_text(raw_brand, max_length=50).upper()
        with col2:
            # Auto-detect general category from sheet name.
            # If brand is already a multi-word label from map_sheet_to_brand
            # (e.g. "NIKE LONG BOTTOMS"), gen_cat should be empty to avoid
            # duplication in the recap label (brand_label = brand + gen_cat).
            if " " in sheet_brand:
                # Multi-word brand already includes category info
                gen_cat_default = ""
            else:
                gen_cat_default = sheet_name.upper().strip()
                if sheet_brand and sheet_brand in gen_cat_default:
                    gen_cat_default = gen_cat_default.replace(sheet_brand, "").strip()
                for prefix in ["BOYS 2-7 ", "GIRLS 2-6X ", "BOYS 4-7 ", "BOYS 8-20 "]:
                    gen_cat_default = gen_cat_default.replace(prefix, "").strip()
                # If nothing left after stripping brand, leave empty (brand alone is fine)
                # Don't fall back to sheet name — that causes "HURLEY HURLEY"
            raw_gen_cat = st.text_input("General Category (e.g. LONG BOTTOMS, TEES)",
                                         value=gen_cat_default, key=f"gencat_{sheet_idx}", max_chars=100)
            gen_cat = sanitize_text(raw_gen_cat, max_length=100).upper()

        # Show detected categories
        st.markdown("**Auto-detected categories:**")
        for cat in categories:
            n_blocks = len(cat["blocks"])
            sr_info = []
            all_refs = []
            for sr_name, sr_data in cat.get("size_ranges", {}).items():
                if sr_data["oh"] > 0 or sr_data["wip"] > 0:
                    sr_info.append(f"{sr_name}: OH={sr_data['oh']:,}")
                for ref in sr_data["refs"]:
                    if ref not in all_refs:
                        all_refs.append(ref)
            sr_str = " | ".join(sr_info) if sr_info else "no data"
            refs_str = ", ".join(sorted(all_refs))
            st.markdown(f"- **{cat['name']}** ({n_blocks} blocks, refs: {refs_str}) — {sr_str}")

        st.success(f"Auto-detected {len(cat_names)} categories with {total_refs} ref#s")

        sheet_configs[sheet_name] = {
            "brand": sheet_brand,
            "general_category": gen_cat,
            "categories": categories,
            "columns": sheet.get("columns"),
        }

# ─── Step 3: Filter Settings ─────────────────────────────────────────────────
st.subheader("Step 3: Filter Settings")
col1, col2 = st.columns(2)
with col1:
    min_units = st.number_input("Minimum units (OH + WIP)", min_value=0, max_value=100000,
                                 value=120, step=10, key="min_units")
with col2:
    use_max = st.checkbox("Set maximum units threshold", key="use_max")
    max_units = None
    if use_max:
        max_units = st.number_input("Maximum units (OH + WIP)", min_value=100,
                                     max_value=1000000, value=12000, step=100, key="max_units")

# ─── Step 4: Generate ─────────────────────────────────────────────────────────
st.subheader("Step 4: Generate Report")

if st.button("Generate ATS Recap Report", type="primary", key="generate_btn"):
    if not check_rate_limit("generate", max_requests=20, window_seconds=3600):
        st.error("Too many generations. Please wait.")
    else:
        with st.spinner("Processing..."):
            try:
                categories_by_sheet = {}
                total_removed = 0

                for sheet_name, config in sheet_configs.items():
                    raw_cats = config["categories"]

                    # Count before
                    before = sum(
                        len([r for r in b["rows"] if r.get("is_label_row", True)])
                        for c in raw_cats for b in c["blocks"]
                    )

                    # Filter
                    filtered_cats = filter_categories(raw_cats, min_units=min_units, max_units=max_units)

                    # Count after
                    after = sum(
                        len([r for r in b["rows"] if r.get("is_label_row", True)])
                        for c in filtered_cats for b in c["blocks"]
                    )
                    total_removed += (before - after)

                    categories_by_sheet[sheet_name] = {
                        "brand": config["brand"],
                        "general_category": config["general_category"],
                        "categories": filtered_cats,
                        "columns": config.get("columns"),
                    }

                logo = parsed.get("logo_image")
                excel_bytes = generate_ats_report(categories_by_sheet,
                                                   title=report_title, report_date=report_date,
                                                   logo_image=logo)

                st.session_state["output_excel"] = excel_bytes
                st.session_state["output_filename"] = f"{report_title}.xlsx"

                c1, c2, c3 = st.columns(3)
                c1.metric("Sheets", len(sheet_configs))
                c2.metric("Categories", sum(len(c["categories"]) for c in categories_by_sheet.values()))
                c3.metric("Styles Removed", total_removed)
                st.success("Report generated successfully!")
                logger.info(f"Report: {report_title} by {user_name}, removed {total_removed}")

            except Exception as e:
                logger.exception("Generation failed")
                st.error(f"Failed to generate report: {e}")

if "output_excel" in st.session_state:
    safe_fn = sanitize_text(st.session_state.get("output_filename", "ATS_RECAP.xlsx"), max_length=100)
    if not safe_fn.endswith('.xlsx'):
        safe_fn += '.xlsx'
    st.download_button("Download ATS Recap Report", data=st.session_state["output_excel"],
                       file_name=safe_fn,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key="download_btn")

with st.sidebar:
    st.markdown("### ATS Recap Report Builder")
    st.markdown("---")
    if st.button("Logout", key="logout_btn"):
        logout()
        st.rerun()
