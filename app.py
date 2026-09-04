"""
GSO/GCTS Tools — web front end
================================
Lets you pick between three tools and run them from a browser instead of
the command line:

  1. GCTS Report generator   (fast, no network, always works hosted)
  2. GSO Model Extractor     (Playwright scraper, public pages, no login)
  3. GSO Product-ID Lookup   (needs a session file generated locally once)

Run locally with:   streamlit run app.py
Free hosting instructions are in README.md.
"""

import os
import io
import json
import subprocess
import sys
import tempfile
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="GSO / GCTS Tools", page_icon="🧰", layout="centered")

# Dark, slightly more dynamic theme for the rest of the page (buttons, cards,
# sidebar) so it doesn't clash with the dark hero banner below.
st.markdown(
    """
    <style>
      .stApp { background: #0b0c0f; }
      section[data-testid="stSidebar"] { background: #111318; }
      .stButton>button {
          background: linear-gradient(135deg, #F9731A, #ff9248);
          color: #0b0c0f; border: none; font-weight: 700;
          transition: transform .15s ease, box-shadow .15s ease;
      }
      .stButton>button:hover { transform: translateY(-1px); box-shadow: 0 6px 18px rgba(249,115,26,.35); }
      div[data-testid="stExpander"] { border: 1px solid #23262e; border-radius: 10px; }
    </style>
    """,
    unsafe_allow_html=True,
)

# Hero banner: particle-text title + spinning globe, both animate in on load.
_hero_html = (Path(__file__).parent / "assets" / "hero.html").read_text()
components.html(_hero_html, height=320, scrolling=False)

# --------------------------------------------------------------------- #
# One-time setup: make sure the Playwright chromium browser is installed.
# On a fresh free-tier host there is no browser binary until this runs.
# --------------------------------------------------------------------- #
@st.cache_resource(show_spinner="Setting up browser engine (first run only)...")
def ensure_playwright_browser():
    try:
        subprocess.run(
            [sys.executable, "-m", "playwright", "install", "--with-deps", "chromium"],
            check=True, capture_output=True, text=True, timeout=600,
        )
        return True
    except Exception as e:
        return f"error: {e}"


st.caption("Pick a tool below. Each one uploads a file, runs the script, and gives you a file back.")

tool = st.sidebar.radio(
    "Choose a tool",
    [
        "1. GCTS Report generator",
        "2. GSO Model Extractor",
        "3. GSO Product-ID Lookup",
        "4. GCTS Expiry Letters",
    ],
)

WORKDIR = Path(tempfile.gettempdir()) / "gso_webapp_runs"
WORKDIR.mkdir(exist_ok=True)

# ======================================================================
# TOOL 1 — GCTS Report generator
# ======================================================================
if tool.startswith("1"):
    st.header("GCTS Lookup → Word report")
    st.write(
        "Upload the **All 35 GCTS Lookup** Excel workbook. Rows whose GCTS number "
        "is highlighted yellow or green are dropped; the rest are grouped by brand "
        "into a formatted Word report."
    )

    xlsx_file = st.file_uploader("Excel file (.xlsx)", type=["xlsx"], key="t1_xlsx")

    if xlsx_file and st.button("Generate report", type="primary"):
        sys.path.insert(0, str(Path(__file__).parent))
        import gcts_report as gr

        with st.spinner("Reading rows and building the document..."):
            tmp_in = WORKDIR / "t1_input.xlsx"
            tmp_in.write_bytes(xlsx_file.getvalue())

            rows = gr.load_rows(str(tmp_in))
            if not rows:
                st.error("No non-highlighted rows found — check the sheet name / highlight colors.")
            else:
                tmp_out = WORKDIR / "GCTS_Report.docx"
                gr.build_report(rows, str(tmp_out))
                st.success(f"Done — {len(rows)} rows across {len(gr.group_by_brand(rows))} brand table(s).")
                st.download_button(
                    "⬇️ Download GCTS_Report.docx",
                    data=tmp_out.read_bytes(),
                    file_name="GCTS_Report.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

# ======================================================================
# TOOL 2 — GSO Model Extractor
# ======================================================================
elif tool.startswith("2"):
    st.header("GSO Model Extractor")
    st.write(
        "Upload an Excel file with a **GSO Product Page** URL column (or a "
        "**Product GCTS** number column). This scrapes each public product page "
        "for its Model Numbers table and Product Specification fields."
    )
    st.warning(
        "⚠️ This runs a real headless browser per row, so it's slow (roughly "
        "1–2 seconds/row) and free hosts will kill requests that run too long. "
        "Use the row limit below to test first — for a full multi-thousand-row "
        "run, do that locally instead of through the browser."
    )

    xlsx_file = st.file_uploader("Excel file (.xlsx)", type=["xlsx"], key="t2_xlsx")
    test_rows = st.number_input(
        "Only process this many rows (0 = no limit — not recommended on a free host)",
        min_value=0, value=5, step=1,
    )

    if xlsx_file and st.button("Run extractor", type="primary"):
        setup = ensure_playwright_browser()
        if setup is not True:
            st.error(f"Could not install the browser engine: {setup}")
        else:
            tmp_in = WORKDIR / "GSO_Batch2_Input.xlsx"
            tmp_in.write_bytes(xlsx_file.getvalue())

            env = os.environ.copy()
            env["WEBAPP_INPUT_FILE"] = str(tmp_in)
            if test_rows > 0:
                env["WEBAPP_TEST_ROWS"] = str(test_rows)

            log_box = st.empty()
            logs = []
            with st.spinner("Scraping pages..."):
                proc = subprocess.Popen(
                    [sys.executable, str(Path(__file__).parent / "gso_model_extractor.py")],
                    cwd=WORKDIR, env=env, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
                )
                for line in proc.stdout:
                    logs.append(line.rstrip())
                    log_box.code("\n".join(logs[-30:]))
                proc.wait()

            out_path = tmp_in.with_name(tmp_in.stem + "_filled.xlsx")
            if proc.returncode == 0 and out_path.exists():
                st.success("Done.")
                st.download_button(
                    "⬇️ Download results",
                    data=out_path.read_bytes(),
                    file_name=out_path.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.error(f"Script exited with code {proc.returncode}. See log above.")

# ======================================================================
# TOOL 3 — GSO Product-ID Lookup
# ======================================================================
elif tool.startswith("3"):
    st.header("GSO Product-ID → GCTS/NB Lookup")
    st.write(
        "This tool needs a **logged-in session** to call the GSO lookup API. "
        "A website can't do the interactive login step for you, so you generate "
        "the session once on your own computer, then upload it here."
    )

    with st.expander("How to get the session file (one-time, on your own computer)"):
        st.markdown(
            "1. Run `python gso_product_id_lookup.py` locally.\n"
            "2. A real browser window opens — log in to GSO normally.\n"
            "3. Press Enter in the terminal once logged in.\n"
            "4. This creates a `gso_state.json` file next to the script.\n"
            "5. Upload that file below. Sessions expire after a while — "
            "regenerate it if lookups start failing."
        )

    input_file = st.file_uploader("Input Excel file (.xlsx)", type=["xlsx"], key="t3_xlsx")
    state_file = st.file_uploader("Session file (gso_state.json)", type=["json"], key="t3_state")
    sheet_name = st.text_input("Sheet name", value="Model Lookup")
    id_col = st.text_input("Product ID column name", value="Product GCTS")
    test_limit = st.number_input(
        "Only process this many new rows (0 = no limit — not recommended on a free host)",
        min_value=0, value=10, step=1,
    )

    if input_file and state_file and st.button("Run lookup", type="primary"):
        try:
            json.loads(state_file.getvalue())
        except Exception:
            st.error("That doesn't look like a valid session JSON file.")
            st.stop()

        tmp_in = WORKDIR / "t3_input.xlsx"
        tmp_in.write_bytes(input_file.getvalue())
        tmp_state = WORKDIR / "gso_state.json"
        tmp_state.write_bytes(state_file.getvalue())
        tmp_out = WORKDIR / "t3_output.xlsx"
        tmp_progress = WORKDIR / "t3_progress.json"

        # patch the sheet/column names the script reads at import time
        script_src = (Path(__file__).parent / "gso_product_id_lookup.py").read_text()
        script_src = script_src.replace(
            'INPUT_SHEET = "Model Lookup"', f'INPUT_SHEET = {sheet_name!r}'
        ).replace(
            'PRODUCT_ID_COLUMN = "Product GCTS"', f'PRODUCT_ID_COLUMN = {id_col!r}'
        )
        run_script = WORKDIR / "_t3_run.py"
        run_script.write_text(script_src)

        env = os.environ.copy()
        env["WEBAPP_INPUT_FILE"] = str(tmp_in)
        env["WEBAPP_OUTPUT_FILE"] = str(tmp_out)
        env["WEBAPP_STATE_FILE"] = str(tmp_state)
        if test_limit > 0:
            env["WEBAPP_TEST_LIMIT"] = str(test_limit)

        # progress file lives next to the script (PROGRESS_FILE is not overridable)
        prog_target = WORKDIR / "gcts_lookup_progress.json"
        if prog_target.exists():
            prog_target.unlink()

        log_box = st.empty()
        logs = []
        with st.spinner("Looking up product IDs..."):
            proc = subprocess.Popen(
                [sys.executable, str(run_script)],
                cwd=WORKDIR, env=env, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
            )
            for line in proc.stdout:
                logs.append(line.rstrip())
                log_box.code("\n".join(logs[-30:]))
            proc.wait()

        if tmp_out.exists():
            st.success("Done — see log above for Found/Not Found counts.")
            st.download_button(
                "⬇️ Download results",
                data=tmp_out.read_bytes(),
                file_name="GSO_merged_with_gcts.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.error(f"Script exited with code {proc.returncode} and produced no output. See log above.")

# ======================================================================
# TOOL 4 — GCTS Expiry Letters (editable contacts, no code editing needed)
# ======================================================================
else:
    import pandas as pd

    st.header("GCTS Expiry Letter Generator")
    st.write(
        "Upload the GCTS Lookup workbook, edit the two tables below to control "
        "who each brand's letter goes to and their email addresses, then generate "
        "one combined Word doc with one letter per contact."
    )

    DATA_DIR = Path(__file__).parent / "data"
    BRAND_CSV = DATA_DIR / "brand_contacts.csv"
    EMAIL_CSV = DATA_DIR / "contact_emails.csv"

    if "brand_contacts_df" not in st.session_state:
        st.session_state.brand_contacts_df = pd.read_csv(BRAND_CSV).fillna("")
    if "contact_emails_df" not in st.session_state:
        st.session_state.contact_emails_df = pd.read_csv(EMAIL_CSV).fillna("")

    xlsx_file = st.file_uploader("GCTS Lookup Excel file (.xlsx)", type=["xlsx"], key="t4_xlsx")

    st.subheader("1. Brand → Contact routing")
    st.caption(
        "Leave 'product_override' blank to route the whole brand to one contact. "
        "Fill it in only for the rare case where different product types under the "
        "same brand go to different people (must exactly match the Product Name column)."
    )
    up_brand = st.file_uploader("...or load a previously saved routing CSV", type=["csv"], key="t4_brand_csv")
    if up_brand is not None:
        st.session_state.brand_contacts_df = pd.read_csv(up_brand).fillna("")
    st.session_state.brand_contacts_df = st.data_editor(
        st.session_state.brand_contacts_df,
        num_rows="dynamic", use_container_width=True, key="t4_brand_editor",
        column_config={
            "brand": st.column_config.TextColumn("Brand", required=True),
            "contact": st.column_config.TextColumn("Contact", required=True),
            "product_override": st.column_config.TextColumn("Product override (optional)"),
        },
    )
    st.download_button(
        "⬇️ Download this routing table as CSV",
        data=st.session_state.brand_contacts_df.to_csv(index=False).encode("utf-8"),
        file_name="brand_contacts.csv", mime="text/csv",
    )

    st.subheader("2. Contact → Email addresses")
    st.caption("One row per email — give a contact multiple rows for multiple addresses.")
    up_email = st.file_uploader("...or load a previously saved emails CSV", type=["csv"], key="t4_email_csv")
    if up_email is not None:
        st.session_state.contact_emails_df = pd.read_csv(up_email).fillna("")
    st.session_state.contact_emails_df = st.data_editor(
        st.session_state.contact_emails_df,
        num_rows="dynamic", use_container_width=True, key="t4_email_editor",
        column_config={
            "contact": st.column_config.TextColumn("Contact", required=True),
            "email": st.column_config.TextColumn("Email", required=True),
        },
    )
    st.download_button(
        "⬇️ Download this email table as CSV",
        data=st.session_state.contact_emails_df.to_csv(index=False).encode("utf-8"),
        file_name="contact_emails.csv", mime="text/csv",
    )

    if xlsx_file and st.button("Generate letters", type="primary"):
        sys.path.insert(0, str(Path(__file__).parent))
        import gso_expiry_letter as gel

        with st.spinner("Building letters..."):
            tmp_in = WORKDIR / "t4_input.xlsx"
            tmp_in.write_bytes(xlsx_file.getvalue())

            rows = gel.load_rows(str(tmp_in))
            if not rows:
                st.error("No non-highlighted rows found — check the sheet name / highlight colors.")
                st.stop()

            brand_contacts = [
                (r["brand"], r["contact"], r.get("product_override", ""))
                for r in st.session_state.brand_contacts_df.to_dict("records")
                if str(r.get("brand", "")).strip() and str(r.get("contact", "")).strip()
            ]
            contact_emails = {}
            for r in st.session_state.contact_emails_df.to_dict("records"):
                c, e = str(r.get("contact", "")).strip(), str(r.get("email", "")).strip()
                if c and e:
                    contact_emails.setdefault(c, []).append(e)

            brand_defaults, product_overrides = gel.build_lookup_dicts(brand_contacts)
            tmp_out = WORKDIR / "GCTS_Expiry_Letters.docx"
            groups = gel.build(rows, brand_defaults, product_overrides, str(tmp_out), contact_emails)

        st.success(f"Done — {len(groups)} letter(s) covering {len(rows)} rows.")
        for contact, contact_rows in groups:
            emails = contact_emails.get(contact, [])
            flag = "" if emails else " — ⚠️ no email on file"
            st.write(f"- **{contact}**: {len(contact_rows)} row(s){flag}")

        st.download_button(
            "⬇️ Download GCTS_Expiry_Letters.docx",
            data=tmp_out.read_bytes(),
            file_name="GCTS_Expiry_Letters.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
