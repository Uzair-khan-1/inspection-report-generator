"""
app.py
------
AI Inspection Report Generator (Streamlit)

Workflow:  Location -> Finding -> Before Photo -> After Photo -> Add -> Next
Then:      Generate Report -> Download Excel

Kept intentionally minimal/fast:
  * No database, no preview pages, no image processing beyond a local resize.
  * One AI call handles ALL pending items' text at once (only when the user
    clicks "Generate Report"), never per-item, to minimize latency.
  * Photos are resized locally with Pillow and placed automatically -- the
    user never touches Excel directly.
  * Location/Finding text is KEPT after "Add item" (many observations share
    the same building/area back-to-back) -- only the photo pickers reset,
    since a fresh pair of photos is needed for every new item.
"""

import os
from datetime import date

import streamlit as st

from ai_service import process_items_with_ai
from excel_processor import generate_report

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "assets", "Template.xlsx")
OUTPUT_PATH = "/tmp/Inspection_Report.xlsx"

st.set_page_config(page_title="AI Inspection Report Generator", page_icon="📋", layout="centered")

if "inspection_items" not in st.session_state:
    st.session_state.inspection_items = []
if "report_bytes" not in st.session_state:
    st.session_state.report_bytes = None
if "uploader_version" not in st.session_state:
    st.session_state.uploader_version = 0  # bump this to reset only the photo pickers

st.title("📋 AI Inspection Report Generator")
st.caption("Location → Finding → Before photo → After photo → **Add**. Repeat, then **Generate Report**.")

# --------------------------------------------------------------------------- #
# Add-item inputs
# (Not wrapped in st.form anymore -- clear_on_submit wiped every field,
# including the text boxes. Now only the photo pickers reset after Add.)
# --------------------------------------------------------------------------- #
col1, col2 = st.columns(2)
location = col1.text_input("Building / Location", key="location_input", placeholder="e.g. room near main gate")
finding = col2.text_input("Finding", key="finding_input", placeholder="e.g. AC leaking water")

col3, col4 = st.columns(2)
before_photo = col3.file_uploader(
    "Before photo", type=["jpg", "jpeg", "png"], key=f"before_uploader_{st.session_state.uploader_version}"
)
after_photo = col4.file_uploader(
    "After photo", type=["jpg", "jpeg", "png"], key=f"after_uploader_{st.session_state.uploader_version}"
)

if st.button("➕ Add item", type="primary", use_container_width=True):
    if not location.strip() or not finding.strip():
        st.warning("Please enter both a Location and a Finding.")
    elif not before_photo or not after_photo:
        st.warning("Please upload both a Before and an After photo.")
    else:
        st.session_state.inspection_items.append(
            {
                "raw_location": location.strip(),
                "raw_finding": finding.strip(),
                "before_bytes": before_photo.getvalue(),
                "after_bytes": after_photo.getvalue(),
            }
        )
        st.session_state.report_bytes = None  # stale, needs regeneration
        st.session_state.uploader_version += 1  # fresh, empty photo pickers next round
        st.rerun()

# --------------------------------------------------------------------------- #
# Pending items list
# --------------------------------------------------------------------------- #
n = len(st.session_state.inspection_items)
if n:
    st.subheader(f"Items added ({n})")
    for i, it in enumerate(st.session_state.inspection_items):
        c1, c2, c3 = st.columns([3, 5, 1])
        c1.markdown(f"**{it['raw_location']}**")
        c2.markdown(it["raw_finding"])
        if c3.button("🗑️", key=f"del_{i}", help="Remove this item"):
            st.session_state.inspection_items.pop(i)
            st.session_state.report_bytes = None
            st.rerun()

    st.divider()
    if st.button("📄 Generate Report", type="primary", use_container_width=True):
        with st.spinner("Cleaning up text and building your Excel report..."):
            # ONE batched AI request for every item's text
            ai_input = [
                {"location": it["raw_location"], "finding": it["raw_finding"]}
                for it in st.session_state.inspection_items
            ]
            cleaned = process_items_with_ai(ai_input)

            report_items = []
            for it, c in zip(st.session_state.inspection_items, cleaned):
                report_items.append(
                    {
                        "location": c["location"],
                        "finding": c["finding"],
                        "before_bytes": it["before_bytes"],
                        "after_bytes": it["after_bytes"],
                        "completed_on": date.today(),
                        "status": "Open",
                    }
                )

            generate_report(report_items, TEMPLATE_PATH, OUTPUT_PATH)
            with open(OUTPUT_PATH, "rb") as f:
                st.session_state.report_bytes = f.read()

    if st.session_state.report_bytes:
        st.success("Report generated!")
        st.download_button(
            "⬇️ Download Excel Report",
            data=st.session_state.report_bytes,
            file_name="Inspection_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
else:
    st.info("No items yet — add your first inspection item above.")
