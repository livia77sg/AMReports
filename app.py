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

st.set_page_config(page_title="Offline Conversions — Automation", page_icon="📊", layout="centered")

st.title("📊 Offline Conversions Automation")
st.caption("Upload the Track360 report and get all 7 conversion reports ready to go.")

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
        """
    )

uploaded = st.file_uploader("Upload the Track360 report (.xlsx)", type=["xlsx"])

if uploaded:
    if st.button("▶️ Process", type="primary", use_container_width=True):
        try:
            with st.spinner("Processing..."):
                result = process_report(
                    report_bytes=uploaded.read(),
                    templates_dir=TEMPLATES_DIR,
                    brand_map_path=BRAND_MAP_PATH,
                )
        except Exception as e:
            st.error(f"❌ Error during processing: {e}")
            st.stop()

        st.success(f"✅ Processed report for date **{result.report_date}**")

        # Summary metrics
        c1, c2, c3 = st.columns(3)
        c1.metric("Reports generated", len(result.output_files))
        c2.metric("Total signups", result.total_signups)
        c3.metric("Total FTD rows", result.total_ftds)

        # Per-output breakdown
        with st.expander("📋 Per-report details"):
            for key, counts in result.per_output_counts.items():
                st.write(
                    f"**{key}** — {counts['signups']} signups, "
                    f"{counts['ftd_rows']} ftd rows"
                )

        # Warnings
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
        )

        # Individual file downloads as fallback
        with st.expander("📄 Or download files one by one"):
            for filename, content in result.output_files.items():
                st.download_button(
                    label=filename,
                    data=content,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=filename,
                )
else:
    st.info("⬆️ Upload a report to get started.")

st.divider()
with st.expander("⚙️ Brand map (brand_map.csv)"):
    if BRAND_MAP_PATH.exists():
        st.caption(f"Location: `{BRAND_MAP_PATH}` — edit this file when new brands appear.")
        with open(BRAND_MAP_PATH) as f:
            st.code(f.read(), language="csv")
    else:
        st.warning(f"Missing `brand_map.csv` at {BRAND_MAP_PATH}")