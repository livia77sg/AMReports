"""Streamlit UI for the offline conversion automation tool.

Run with:
    streamlit run app.py
"""

import io
import zipfile
from pathlib import Path

import streamlit as st

from processor import process_report

BASE_DIR = Path(__file__).parent
TEMPLATES_DIR = BASE_DIR / "templates"
BRAND_MAP_PATH = BASE_DIR / "brand_map.csv"
BRAND_VALUES_PATH = BASE_DIR / "brand_values.csv"

st.set_page_config(page_title="Offline Conversions — Automation", page_icon="📊", layout="centered")

st.title("📊 Offline Conversions Automation")
st.caption("Upload the Track360 report and get all conversion reports ready to go.")

# Initialize session state — keeps the processing result around so download
# buttons don't trigger a full rerun and lose the data.
if "result" not in st.session_state:
    st.session_state.result = None
if "uploaded_filename" not in st.session_state:
    st.session_state.uploaded_filename = None
# Counter used as the file_uploader key — bumping it on reset forces Streamlit
# to mount a fresh widget, which is the only way to clear the upload field.
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

with st.expander("ℹ️ How it works"):
    st.markdown(
        """
        1. Download yesterday's report from **Track360**.
        2. Upload it below.
        3. Click **Process**.
        4. Download the ZIP with all the generated reports.

        **What the app does:**
        - Filters out rows where `Cid = 'na'`.
        - Strips the `m_` prefix from Cid for Bing.
        - Identifies bingo / casino / sport brands using `brand_map.csv`.
        - Collapses casino brands in Sport/Bingo into a single `offline - ftdcasino` row.
        - Splits bingo and sport brands into two rows (brand name + `offline - ftds`).
        - Sets time **23:55** for signups and **23:58** for ftds.
        - Rounds revenue to a whole number.
        - Skips outputs with no data (no empty files).
        """
    )

uploaded = st.file_uploader(
    "Upload the Track360 report (.xlsx)",
    type=["xlsx"],
    key=f"uploader_{st.session_state.uploader_key}",
)

# If a new file is uploaded, clear stale result so it doesn't show old data
if uploaded and uploaded.name != st.session_state.uploaded_filename:
    st.session_state.result = None
    st.session_state.uploaded_filename = uploaded.name

col1, col2 = st.columns([3, 1])
with col1:
    process_clicked = st.button(
        "▶️ Process",
        type="primary",
        use_container_width=True,
        disabled=uploaded is None,
    )
with col2:
    reset_clicked = st.button(
        "🔄 Reset",
        use_container_width=True,
        help="Clear the current result and start over.",
    )

if reset_clicked:
    st.session_state.result = None
    st.session_state.uploaded_filename = None
    st.session_state.uploader_key += 1  # force the file_uploader to remount empty
    st.rerun()

if process_clicked and uploaded:
    try:
        with st.spinner("Processing..."):
            result = process_report(
                report_bytes=uploaded.read(),
                templates_dir=TEMPLATES_DIR,
                brand_map_path=BRAND_MAP_PATH,
                brand_values_path=BRAND_VALUES_PATH,
            )
        # Cache the result so downloads don't re-trigger processing
        st.session_state.result = result
    except Exception as e:
        st.session_state.result = None
        st.error(f"❌ Error during processing: {e}")
        st.stop()

# Render results (from session_state so they persist across reruns) ---------
result = st.session_state.result
if result is not None:
    st.success(f"✅ Processed report for date **{result.report_date}**")

    c1, c2, c3 = st.columns(3)
    c1.metric("Reports generated", len(result.output_files))
    c2.metric("Total signups", result.total_signups)
    c3.metric("Total FTD rows", result.total_ftds)

    with st.expander("📋 Per-report details"):
        for key, counts in result.per_output_counts.items():
            had_data = counts["signups"] + counts["ftd_rows"] > 0
            status = "" if had_data else " (skipped — no data)"
            st.write(
                f"**{key}** — {counts['signups']} signups, "
                f"{counts['ftd_rows']} ftd rows{status}"
            )

    if result.warnings:
        with st.expander(f"⚠️ Warnings ({len(result.warnings)})", expanded=True):
            for w in result.warnings:
                st.warning(w)

    if result.unmapped_brands:
        with st.expander(f"🏷️ Unmapped brands ({len(result.unmapped_brands)})"):
            st.write(
                "These brands weren't found in `brand_map.csv` and were processed "
                "with the default rule (lowercase + suffix removal). "
                "Add them to the CSV for full control."
            )
            st.code("\n".join(result.unmapped_brands))

    if getattr(result, "missing_cpa_brands", None):
        with st.expander(
            f"💰 Brands missing CPA value ({len(getattr(result, 'missing_cpa_brands', []))})",
            expanded=True,
        ):
            st.write(
                "These brands are mapped in `brand_map.csv` but **don't have a CPA value** "
                "in `brand_values.csv` — so their rows were **skipped**. "
                "Add them to `brand_values.csv` with the correct value to include them."
            )
            # Group by vertical for clarity
            by_vertical = {}
            for raw, vert in getattr(result, "missing_cpa_brands", []):
                by_vertical.setdefault(vert, []).append(raw)
            for vert in sorted(by_vertical):
                st.write(f"**{vert.upper()}:**")
                st.code("\n".join(sorted(by_vertical[vert])))

    # ZIP download
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for filename, content in result.output_files.items():
            zf.writestr(filename, content)
    zip_buf.seek(0)

    st.download_button(
        label=f"📥 Download ZIP ({len(result.output_files)} files)",
        data=zip_buf,
        file_name=f"offline_conversions_{result.report_date}.zip",
        mime="application/zip",
        use_container_width=True,
        key="zip_download",
    )

    with st.expander("📄 Or download files one by one"):
        for filename, content in result.output_files.items():
            st.download_button(
                label=filename,
                data=content,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"individual_{filename}",
            )
elif uploaded is None:
    st.info("⬆️ Upload a report to get started.")

st.divider()
with st.expander("⚙️ Brand map (brand_map.csv)"):
    if BRAND_MAP_PATH.exists():
        st.caption(
            f"Location: `{BRAND_MAP_PATH}` — edit this file when new brands appear. "
            "The optional `output_key` column lets you override a brand name for a specific output "
            "(e.g. `LottoGo-Casino-GB` becomes `Lottogo` in UK1 but `lottogo` everywhere else)."
        )
        with open(BRAND_MAP_PATH) as f:
            st.code(f.read(), language="csv")
    else:
        st.warning(f"Missing `brand_map.csv` at {BRAND_MAP_PATH}")

with st.expander("💰 Brand values / CPA (brand_values.csv)"):
    if BRAND_VALUES_PATH.exists():
        st.caption(
            f"Location: `{BRAND_VALUES_PATH}` — per-brand fixed CPA values "
            "for the HW (Home Warranty) and CW (Car Warranty) verticals. "
            "Edit when rates change."
        )
        with open(BRAND_VALUES_PATH) as f:
            st.code(f.read(), language="csv")
    else:
        st.warning(f"Missing `brand_values.csv` at {BRAND_VALUES_PATH}")