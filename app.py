"""
ATS Recap Report Builder — Main Application
Haddad Brands Value Channel Sales

Upload raw ATS Excel -> Clean, format, and generate buyer-ready recap report.

Security: Input validation, rate limiting, no stack trace exposure,
          file validation, formula injection prevention.
"""

import logging
from datetime import date

import streamlit as st

from utils.auth import login, check_password, logout
from utils.ats_parser import (
    parse_ats_file, filter_blocks, group_blocks_by_category,
    detect_brand,
)
from utils.excel_generator import generate_ats_report
from utils.security import (
    sanitize_text, sanitize_for_excel, validate_xlsx_file,
    validate_category_mapping, check_rate_limit,
)

# ─── Logging (server-side only, never exposed to users) ─────────────────────
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(name)s: %(message)s',
    handlers=[logging.StreamHandler()],
)
logger = logging.getLogger(__name__)

st.set_page_config(
    page_title="ATS Recap Report Builder",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─── Global CSS (static only — no user data in HTML) ────────────────────────
# All CSS is static content, safe to use unsafe_allow_html
st.markdown("""
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}

body { background: #F0F4F8; }

.login-card {
    background: white;
    border-radius: 16px;
    padding: 48px 40px 40px;
    box-shadow: 0 8px 32px rgba(27,43,75,0.12);
    max-width: 420px;
    margin: 60px auto 0;
}
.login-logo { font-size: 2.8rem; text-align: center; margin-bottom: 6px; }
.login-title { font-size: 1.5rem; font-weight: 700; color: #1B2B4B; text-align: center; margin-bottom: 4px; }
.login-subtitle { font-size: 0.9rem; color: #6B7C99; text-align: center; margin-bottom: 32px; }

.stTextInput > div > div > input {
    border-radius: 8px !important;
    border: 1.5px solid #D0D7E5 !important;
    padding: 12px 14px !important;
}
.stTextInput > div > div > input:focus {
    border-color: #00B0F0 !important;
    box-shadow: 0 0 0 3px rgba(0,176,240,0.15) !important;
}
.stButton > button {
    width: 100%;
    background: linear-gradient(135deg, #00B0F0, #0080C0) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 12px !important;
    font-weight: 600 !important;
}
.stButton > button:hover { opacity: 0.9; }
</style>
""", unsafe_allow_html=True)


# ─── Login ───────────────────────────────────────────────────────────────────
if not check_password():
    # Static HTML only — no user data injected
    st.markdown("""
    <div class="login-card">
      <div class="login-logo">📊</div>
      <div class="login-title">ATS Recap Report Builder</div>
      <div class="login-subtitle">Value Channel Sales</div>
    </div>
    """, unsafe_allow_html=True)

    col_l, col_m, col_r = st.columns([1, 2, 1])
    with col_m:
        st.write("")
        name = st.text_input("Your Name", placeholder="e.g. Jacob H.", key="login_name",
                             max_chars=50)
        password = st.text_input("Password", type="password", placeholder="Enter team password",
                                  key="login_pw", max_chars=100)
        if st.button("Sign In", key="login_btn"):
            if not name.strip():
                st.error("Please enter your name.")
            elif not login(name, password):
                st.error("Incorrect password or too many attempts. Try again.")
            else:
                st.rerun()
    st.stop()


# ─── Main App (Authenticated) ───────────────────────────────────────────────

# Display user name safely (no unsafe_allow_html with user data)
user_name = st.session_state.get('user_name', '')
st.title("ATS Recap Report Builder")
st.caption(f"Upload → Configure → Download  |  Welcome, {user_name}")


# ─── Step 1: Upload ─────────────────────────────────────────────────────────

st.subheader("Step 1: Upload Raw ATS File")

uploaded_file = st.file_uploader(
    "Upload your raw ATS Excel file (.xlsx)",
    type=["xlsx"],
    key="ats_upload",
    help="Upload the Excel file with raw ATS data pasted onto sheets.",
)

if uploaded_file is None:
    st.info("Upload your raw ATS Excel file to get started.")
    st.stop()


# ─── Validate and Parse ────────────────────────────────────────────────────

file_bytes = uploaded_file.read()

# File validation (size, MIME type, zip bomb check)
is_valid, error_msg = validate_xlsx_file(file_bytes, uploaded_file.name)
if not is_valid:
    st.error(f"File validation failed: {error_msg}")
    logger.warning(f"File validation failed for {uploaded_file.name}: {error_msg}")
    st.stop()

# Rate limit file uploads
if not check_rate_limit("upload", max_requests=30, window_seconds=3600):
    st.error("Too many uploads. Please wait before trying again.")
    st.stop()


# Parse with short-lived cache (5 min TTL, no persistent sensitive data)
@st.cache_data(ttl=300, show_spinner="Parsing file...")
def cached_parse(file_bytes_input):
    return parse_ats_file(file_bytes_input)


try:
    parsed = cached_parse(file_bytes)
except Exception:
    logger.exception("Failed to parse uploaded file")
    st.error("Could not parse the uploaded file. Please verify it is a valid ATS Excel file.")
    st.stop()

if not parsed["sheets"]:
    st.error("No data found in the uploaded file. Make sure it contains ATS data.")
    st.stop()


# ─── Step 2: Configure Sheets ───────────────────────────────────────────────

st.subheader("Step 2: Configure Sheets & Categories")

# Report title (sanitized)
raw_title = st.text_input(
    "Report Title (for RECAP sheet header)",
    value=uploaded_file.name.replace(".xlsx", "").replace(".xls", "").upper(),
    key="report_title",
    max_chars=200,
)
report_title = sanitize_for_excel(sanitize_text(raw_title, max_length=200))

report_date = st.date_input("Report Date", value=date.today(), key="report_date")

st.markdown("---")

# For each sheet, show detected ref#s and let user configure
sheet_configs = {}

for sheet_idx, sheet in enumerate(parsed["sheets"]):
    sheet_name = sheet["name"]
    brand = sheet["brand"]
    blocks = sheet["blocks"]
    ref_nums = sheet["all_ref_nums"]

    if not blocks:
        continue

    with st.expander(f"📋 {sheet_name} ({len(ref_nums)} ref#s detected)", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            detected_brand = brand if brand else "UNKNOWN"
            raw_brand = st.text_input(
                "Brand", value=detected_brand, key=f"brand_{sheet_idx}",
                max_chars=50,
            )
            sheet_brand = sanitize_text(raw_brand, max_length=50).upper()

        with col2:
            gen_cat_default = sheet_name.upper()
            if sheet_brand and sheet_brand in gen_cat_default:
                gen_cat_default = gen_cat_default.replace(sheet_brand, "").strip()
            for prefix in ["BOYS 2-7 ", "GIRLS 2-6X ", "BOYS 4-7 ", "BOYS 8-20 "]:
                gen_cat_default = gen_cat_default.replace(prefix, "").strip()
            raw_gen_cat = st.text_input(
                "General Category (e.g. LONG BOTTOMS, TEES)",
                value=gen_cat_default, key=f"gencat_{sheet_idx}",
                max_chars=100,
            )
            gen_cat = sanitize_text(raw_gen_cat, max_length=100).upper()

        # Display ref#s safely (using st.code, not unsafe HTML)
        st.markdown(f"**Detected Ref#s:** `{', '.join(ref_nums[:100])}`")
        if len(ref_nums) > 100:
            st.caption(f"... and {len(ref_nums) - 100} more")

        st.markdown("**Define categories** — one per line: `CATEGORY NAME: REF1, REF2, REF3`")
        st.code("BURPEE JOGGER: F610, F651\nTHERMA PANT: J785, N271", language=None)

        default_mapping = st.session_state.get(f"catmap_{sheet_idx}", "")
        if not default_mapping:
            default_mapping = "\n".join(f"CATEGORY_{ref}: {ref}" for ref in ref_nums)

        cat_mapping_text = st.text_area(
            "Category → Ref# Mapping",
            value=default_mapping,
            height=200,
            key=f"catmap_input_{sheet_idx}",
            help="One category per line. Format: CATEGORY NAME: REF1, REF2, REF3",
        )

        # Validate the mapping with security checks
        is_valid_map, map_result = validate_category_mapping(cat_mapping_text)
        if is_valid_map:
            category_map = map_result
            mapped_refs = set()
            for refs in category_map.values():
                mapped_refs.update(refs)

            unmapped = [r for r in ref_nums if r.upper() not in mapped_refs]
            if unmapped:
                st.warning(f"Unmapped ref#s: {', '.join(unmapped[:50])}")

            st.success(f"{len(category_map)} categories defined covering "
                      f"{len(mapped_refs)}/{len(ref_nums)} ref#s")
        else:
            category_map = {}
            if cat_mapping_text.strip():
                st.warning(f"Mapping issue: {map_result}")

        sheet_configs[sheet_name] = {
            "brand": sheet_brand,
            "general_category": gen_cat,
            "category_map": category_map,
            "blocks": blocks,
        }


# ─── Step 3: Thresholds ─────────────────────────────────────────────────────

st.subheader("Step 3: Filter Settings")

col1, col2 = st.columns(2)
with col1:
    min_units = st.number_input(
        "Minimum units (OH + WIP) — remove styles below this",
        min_value=0, max_value=100000, value=120, step=10,
        key="min_units",
    )
with col2:
    use_max = st.checkbox("Also set a maximum units threshold", key="use_max")
    max_units = None
    if use_max:
        max_units = st.number_input(
            "Maximum units (OH + WIP) — remove styles above this",
            min_value=100, max_value=1000000, value=12000, step=100,
            key="max_units",
        )


# ─── Step 4: Generate ───────────────────────────────────────────────────────

st.subheader("Step 4: Generate Report")

if st.button("Generate ATS Recap Report", type="primary", key="generate_btn"):
    # Rate limit generation
    if not check_rate_limit("generate", max_requests=20, window_seconds=3600):
        st.error("Too many report generations. Please wait before trying again.")
    else:
        with st.spinner("Processing..."):
            try:
                categories_by_sheet = {}
                total_blocks_before = 0
                total_blocks_after = 0
                total_styles_removed = 0

                for sheet_name, config in sheet_configs.items():
                    blocks = config["blocks"]
                    category_map = config["category_map"]
                    total_blocks_before += len(blocks)

                    filtered = filter_blocks(blocks, min_units=min_units, max_units=max_units)
                    total_blocks_after += len(filtered)

                    original_packs = sum(
                        len([r for r in b["rows"] if r.get("is_label_row", True)])
                        for b in blocks
                    )
                    filtered_packs = sum(
                        len([r for r in b["rows"] if r.get("is_label_row", True)])
                        for b in filtered
                    )
                    total_styles_removed += (original_packs - filtered_packs)

                    categories = group_blocks_by_category(filtered, category_map)

                    categories_by_sheet[sheet_name] = {
                        "brand": config["brand"],
                        "general_category": config["general_category"],
                        "categories": categories,
                    }

                # Generate Excel (title already sanitized)
                excel_bytes = generate_ats_report(
                    categories_by_sheet,
                    title=report_title,
                    report_date=report_date,
                )

                st.session_state["output_excel"] = excel_bytes
                st.session_state["output_filename"] = f"{report_title}.xlsx"

                # Show summary using safe Streamlit components (no unsafe HTML with dynamic data)
                c1, c2, c3 = st.columns(3)
                c1.metric("Sheets", len(sheet_configs))
                c2.metric("Ref# Blocks", total_blocks_after)
                c3.metric("Styles Removed", total_styles_removed)

                st.success("Report generated successfully!")
                logger.info(f"Report generated: {report_title} by {user_name} "
                           f"({total_blocks_after} blocks, {total_styles_removed} removed)")

            except Exception:
                logger.exception("Report generation failed")
                st.error("Failed to generate report. Please verify your file and settings, then try again.")

# Download button
if "output_excel" in st.session_state:
    # Sanitize filename
    safe_filename = sanitize_text(
        st.session_state.get("output_filename", "ATS_RECAP_REPORT.xlsx"),
        max_length=100
    )
    if not safe_filename.endswith('.xlsx'):
        safe_filename += '.xlsx'

    st.download_button(
        label="Download ATS Recap Report",
        data=st.session_state["output_excel"],
        file_name=safe_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_btn",
    )

# ─── Sidebar ────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ATS Recap Report Builder")
    st.markdown("---")
    if st.button("Logout", key="logout_btn"):
        logout()
        st.rerun()
