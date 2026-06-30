"""
Conference Soft Slide Generator
Streamlit app for generating event soft slides from a designer template +
mask, session details, and speaker info/photos.

Run with: streamlit run app.py
"""

import io
import os
import streamlit as st
from PIL import Image

from core.mask_parser import load_slot_shape
from core.layout_engine import get_layout, MAX_SUPPORTED, sort_speakers_moderator_first
from core.photo_processor import process_photo
from core.slide_compositor import compose_slide, SpeakerInfo
from core.speaker_sheet import read_speaker_sheet, DEFAULT_ROLE_LABEL
from core.name_matcher import match_photos_to_speakers
from core.pptx_compositor import new_presentation, build_pptx_slide, save_pptx

APP_DIR = os.path.dirname(os.path.abspath(__file__))

st.set_page_config(page_title="Soft Slide Generator", layout="wide")

# ---------------------------------------------------------------------------
# Session state setup
# ---------------------------------------------------------------------------
if "processed_speakers" not in st.session_state:
    st.session_state.processed_speakers = None   # list of dicts after "Process Photos"
if "final_slide" not in st.session_state:
    st.session_state.final_slide = None
if "final_pptx_bytes" not in st.session_state:
    st.session_state.final_pptx_bytes = None

st.title("🖥️ Conference Soft Slide Generator")
st.caption("Upload a template + mask once per event. Generate as many session slides as you need.")

# ---------------------------------------------------------------------------
# STEP 1: Template + Mask (event-level, done once per event/theme)
# ---------------------------------------------------------------------------
st.header("1. Event Template & Photo Shape")

col1, col2 = st.columns(2)
with col1:
    template_file = st.file_uploader(
        "Base template (PNG) — includes logo/branding, empty canvas below",
        type=["png"],
    )
with col2:
    mask_file = st.file_uploader(
        "Photo placeholder mask (PNG) — the shape designers want photos cropped to "
        "(e.g. circle, hexagon). Draw ONE shape in white/opaque on a transparent "
        "or dark background.",
        type=["png"],
    )

template_img, slot_shape = None, None
if template_file:
    template_img = Image.open(template_file)
    st.image(template_img, caption=f"Template preview ({template_img.width}x{template_img.height}px)", width=500)

if mask_file:
    # load_slot_shape expects a file path, so write the upload to a temp file
    import tempfile
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        tmp.write(mask_file.getvalue())
        tmp_path = tmp.name
    try:
        slot_shape = load_slot_shape(tmp_path)
        if slot_shape.warning:
            st.warning(slot_shape.warning)
        st.success(f"Photo shape detected: {slot_shape.width}x{slot_shape.height}px, "
                    f"aspect ratio {slot_shape.aspect_ratio:.2f}")
        st.image(slot_shape.alpha_mask, caption="Extracted slot shape", width=150)
    except Exception as e:
        st.error(f"Could not read mask file: {e}")
        slot_shape = None

st.divider()

# ---------------------------------------------------------------------------
# STEP 2: Session details
# ---------------------------------------------------------------------------
st.header("2. Session Details")

c1, c2, c3 = st.columns(3)
with c1:
    session_name = st.text_input("Session name", placeholder="e.g. AI for India 2030 - Star Panel")
with c2:
    hall_name = st.text_input("Hall name", placeholder="e.g. Hall 1 / Bangalore Palace")
with c3:
    date_str = st.text_input("Date", placeholder="e.g. 19th November, 2026")

st.divider()

# ---------------------------------------------------------------------------
# STEP 3: Speakers (bulk Excel + bulk photo upload, with fuzzy matching)
# ---------------------------------------------------------------------------
st.header("3. Speakers")

entry_mode = st.radio(
    "How would you like to enter speaker details?",
    ["Bulk upload (Excel + photos)", "Manual entry (type each speaker)"],
    horizontal=True,
)

speaker_inputs = []  # final list of dicts: name, title, company, role, photo_bytes

if entry_mode == "Bulk upload (Excel + photos)":
    template_path = os.path.join(APP_DIR, "assets", "speaker_list_template.xlsx")
    with open(template_path, "rb") as f:
        st.download_button(
            "⬇️ Download Excel template", data=f.read(),
            file_name="speaker_list_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    excel_file = st.file_uploader(
        "Upload filled speaker list (.xlsx)", type=["xlsx"], key="excel_upload"
    )
    photo_files = st.file_uploader(
        "Bulk upload speaker photos — filename should contain the speaker's name "
        "in some form (e.g. 'ravi_singh.jpg', 'Singh Ravi.png' — order/case/separators don't matter)",
        type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="bulk_photos",
    )

    speaker_rows = None
    if excel_file is not None:
        try:
            speaker_rows = read_speaker_sheet(excel_file)
            st.success(f"Loaded {len(speaker_rows)} speaker(s) from the sheet.")
        except ValueError as e:
            st.error(str(e))

    if speaker_rows and photo_files:
        names = [r["name"] for r in speaker_rows]
        filenames = [pf.name for pf in photo_files]
        match_results = match_photos_to_speakers(names, filenames)
        filename_to_file = {pf.name: pf for pf in photo_files}

        st.subheader("Review photo matches")
        st.caption("Auto-matched by filename similarity. Fix any incorrect or missing matches below before continuing.")

        all_filenames_options = ["(no photo)"] + filenames
        final_matches = []
        for i, (row, match) in enumerate(zip(speaker_rows, match_results)):
            cols = st.columns([3, 2, 2])
            cols[0].markdown(f"**{row['name']}**" + (f" _({row['role']})_" if row.get("role") else ""))

            default_choice = match.matched_filename if match.matched_filename in filenames else "(no photo)"
            chosen = cols[1].selectbox(
                "Matched photo", options=all_filenames_options,
                index=all_filenames_options.index(default_choice),
                key=f"match_select_{i}",
                label_visibility="collapsed",
            )
            if chosen != "(no photo)":
                cols[2].image(filename_to_file[chosen], width=70)
                if match.needs_review and chosen == match.matched_filename:
                    cols[2].caption("⚠️ low-confidence auto-match — please verify")
            else:
                cols[2].caption("No photo assigned")

            final_matches.append({**row, "photo_file": filename_to_file.get(chosen)})

        speaker_inputs = final_matches

        unmatched_photos = [f for f in filenames if f not in [m["photo_file"].name for m in final_matches if m["photo_file"]]]
        if unmatched_photos:
            st.info(f"{len(unmatched_photos)} uploaded photo(s) were not assigned to any speaker: "
                     + ", ".join(unmatched_photos))

else:
    if "num_speakers" not in st.session_state:
        st.session_state.num_speakers = 3

    num_speakers = st.number_input(
        f"Number of speakers (max {MAX_SUPPORTED} supported)",
        min_value=1, max_value=MAX_SUPPORTED, value=st.session_state.num_speakers, step=1,
    )
    st.session_state.num_speakers = num_speakers

    for i in range(num_speakers):
        with st.expander(f"Speaker {i+1}", expanded=(i < 3)):
            cols = st.columns([2, 2, 2, 2, 1])
            name = cols[0].text_input("Name", key=f"name_{i}")
            title = cols[1].text_input("Title", key=f"title_{i}")
            company = cols[2].text_input("Company", key=f"company_{i}")
            role = cols[3].text_input(
                "Role label (optional)", key=f"role_{i}",
                placeholder="e.g. Moderator, Chief Guest — blank = Speaker",
            )
            photo = cols[4].file_uploader("Photo", type=["jpg", "jpeg", "png"], key=f"photo_{i}")
            speaker_inputs.append({
                "name": name, "title": title, "company": company,
                "role": role.strip() or DEFAULT_ROLE_LABEL, "photo_file": photo,
            })

st.divider()

# ---------------------------------------------------------------------------
# STEP 4: Process photos -> review grid
# ---------------------------------------------------------------------------
st.header("4. Process & Review Photos")

if st.button("🔄 Process Photos", type="primary", disabled=not speaker_inputs):
    missing_photos = [s for s in speaker_inputs if s["photo_file"] is None]
    if missing_photos:
        names_list = ", ".join(s["name"] for s in missing_photos)
        st.error(f"These speaker(s) have no photo assigned: {names_list}. "
                  "Please assign a photo for every speaker before processing.")
    else:
        # Default order: moderator/chair first (matches reference slide
        # convention), preserving relative order otherwise. Staff can
        # still manually reorder afterward using the controls below.
        ordered_inputs = sort_speakers_moderator_first(speaker_inputs)
        processed = []
        for s in ordered_inputs:
            raw_img = Image.open(s["photo_file"])
            result = process_photo(raw_img)
            processed.append({**s, "processed_image": result.image,
                               "face_detected": result.face_detected, "note": result.note})
        st.session_state.processed_speakers = processed
        st.session_state.final_slide = None  # reset downstream state
        st.session_state.final_pptx_bytes = None

if st.session_state.processed_speakers:
    st.subheader("Review processed photos before generating the final slide")
    st.caption(
        "Default order places Moderator/Chair first. Use the ↑ / ↓ buttons "
        "to rearrange speakers into your preferred order before generating."
    )

    speakers_list = st.session_state.processed_speakers
    needs_attention = False

    review_cols = st.columns(min(len(speakers_list), 5))
    for i, sp in enumerate(speakers_list):
        col = review_cols[i % len(review_cols)]
        with col:
            st.image(sp["processed_image"], caption=sp["name"] or f"Speaker {i+1}", width=150)
            if not sp["face_detected"]:
                st.warning("⚠️ " + sp["note"])
                needs_attention = True

            btn_cols = st.columns(2)
            if btn_cols[0].button("↑ Up", key=f"up_{i}", disabled=(i == 0), use_container_width=True):
                speakers_list[i - 1], speakers_list[i] = speakers_list[i], speakers_list[i - 1]
                st.session_state.processed_speakers = speakers_list
                st.session_state.final_slide = None
                st.session_state.final_pptx_bytes = None
                st.rerun()
            if btn_cols[1].button("↓ Down", key=f"down_{i}", disabled=(i == len(speakers_list) - 1), use_container_width=True):
                speakers_list[i + 1], speakers_list[i] = speakers_list[i], speakers_list[i + 1]
                st.session_state.processed_speakers = speakers_list
                st.session_state.final_slide = None
                st.session_state.final_pptx_bytes = None
                st.rerun()

    if needs_attention:
        st.info(
            "Some photos couldn't be auto-cropped confidently. You can still proceed, "
            "but consider re-uploading a clearer photo for flagged speakers above "
            "(crop the photo close to the face yourself before uploading) and re-process."
        )

st.divider()

# ---------------------------------------------------------------------------
# STEP 5: Generate final slide
# ---------------------------------------------------------------------------
st.header("5. Generate Final Slide")

ready = template_img is not None and slot_shape is not None and st.session_state.processed_speakers is not None

if not ready:
    st.info("Complete steps 1-4 above (template, mask, session details, and processed photos) to generate the slide.")

output_format = st.radio(
    "Output format",
    ["PowerPoint (.pptx) — editable, recommended", "PNG image — flat, view-only"],
    horizontal=True,
)

if st.button("✨ Generate Slide", disabled=not ready, type="primary"):
    try:
        layout_slots = get_layout(len(st.session_state.processed_speakers))

        if output_format.startswith("PowerPoint"):
            speaker_dicts = [
                {
                    "name": sp["name"], "title": sp["title"], "company": sp["company"],
                    "role": sp.get("role", DEFAULT_ROLE_LABEL), "photo": sp["processed_image"],
                }
                for sp in st.session_state.processed_speakers
            ]
            prs = new_presentation()
            build_pptx_slide(
                prs, template=template_img, slot_shape=slot_shape, slots=layout_slots,
                speakers=speaker_dicts, session_name=session_name,
                hall_name=hall_name, date_str=date_str,
            )
            pptx_buf = io.BytesIO()
            prs.save(pptx_buf)
            st.session_state.final_pptx_bytes = pptx_buf.getvalue()
            st.session_state.final_slide = None
        else:
            speaker_infos = [
                SpeakerInfo(
                    name=sp["name"], title=sp["title"], company=sp["company"],
                    photo=sp["processed_image"], role_label=sp.get("role", DEFAULT_ROLE_LABEL),
                )
                for sp in st.session_state.processed_speakers
            ]
            final = compose_slide(
                template=template_img, slot_shape=slot_shape, slots=layout_slots,
                speakers=speaker_infos, session_name=session_name,
                hall_name=hall_name, date_str=date_str,
            )
            st.session_state.final_slide = final
            st.session_state.final_pptx_bytes = None
    except ValueError as e:
        st.error(str(e))

if st.session_state.final_pptx_bytes is not None:
    st.success("PowerPoint generated — every text box and photo is editable in PowerPoint.")
    st.download_button(
        "⬇️ Download Slide (PPTX)", data=st.session_state.final_pptx_bytes,
        file_name=f"{(session_name or 'soft_slide').replace(' ', '_')}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )

if st.session_state.final_slide is not None:
    st.subheader("Final Slide (PNG preview)")
    st.image(st.session_state.final_slide, use_container_width=True)

    buf = io.BytesIO()
    st.session_state.final_slide.convert("RGB").save(buf, format="PNG")
    st.download_button(
        "⬇️ Download Slide (PNG)", data=buf.getvalue(),
        file_name=f"{(session_name or 'soft_slide').replace(' ', '_')}.png",
        mime="image/png",
    )
