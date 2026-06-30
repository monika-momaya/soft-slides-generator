# Conference Soft Slide Generator (v2 - Editable PowerPoint Output)

Generates ready-to-use conference slides from:
- a designer's base template (PNG, branding pre-baked in)
- a designer's mask PNG (one reusable photo-slot shape, e.g. circle/hexagon)
- session details + speaker info/photos entered by event staff

**Primary output is an editable .pptx file** — every text box (session
title, speaker names/titles/companies, role labels) is a real, editable
PowerPoint object your team can tweak live during an event (fix a typo,
swap a name, adjust a position) without regenerating anything. Speaker
photos are inserted as picture shapes, cropped to the designer's exact
mask shape. A flat PNG export is also available as a secondary option.

## What's new in this version (round 2 fixes)

- **Fixed: full slide width now used in all layouts.** Earlier 7/8-speaker
  layouts only spread across roughly two-thirds of the slide, leaving a
  large empty gap on the right. All layouts now use the full slide width.
- **Fixed: consistent speaker name sizing.** All speaker names on a slide
  now share ONE font size by default (sized to fit whichever name is
  longest), instead of each name being sized independently. This means
  selecting all the name text boxes in PowerPoint and changing the font
  size actually behaves as expected.
- **Moderator/Chair now appears first.** Speakers with a Role of
  "Moderator", "Chair", "Co-Moderator", or "Co-Chair" are automatically
  sorted to the front of the lineup by default (leftmost/topmost slot),
  matching the convention in reference event slides. This is a sorting
  step applied before layout, not a layout-level assumption — it doesn't
  reintroduce the earlier position-based labeling bug.
- **Manual speaker reordering.** In Step 4 (Process & Review Photos), each
  speaker now has ↑ / ↓ buttons to move them earlier or later in the
  lineup, overriding the default moderator-first order if needed.
- **Tighter spacing**: the gap between a speaker's name and their title/
  company line is reduced, and the gap between the session title and the
  hall/date line is now close to single-line spacing (was noticeably too
  large before).
- **Smart title/company text layout**: short title + company combine onto
  one line (e.g. "Founder, Infosys"). Longer text wraps using simple
  width-based wrapping (breaks wherever it hits the box edge, not
  phrase-aware) rather than shrinking to an illegibly small font.

## What's new in this version (PPTX rebuild)

- **PowerPoint (.pptx) output is now the primary format**, generated with
  real editable text boxes and picture shapes — not a flattened image.
  PNG export is still available as a secondary option.
- **Fixed a real bug**: role/moderator labels are now tied strictly to
  each speaker's own data, never to their position in the speaker list.
  (Earlier versions hardcoded "MODERATOR" onto whichever speaker happened
  to be first in a layout slot — this has been removed entirely.)
- **Open-ended role labels.** The old `Moderator (Y/N)` column is replaced
  with a free-text `Role` column. Write anything descriptive — "Moderator",
  "Chief Guest", "Keynote Speaker", "Event Inauguration by" — and it's
  used verbatim as the label shown above that speaker's photo. Leave it
  blank and the speaker is treated as a regular Speaker/Panelist (no
  label shown, matching the convention in reference event slides).
- **Widened high-density layouts** (7-8 speakers) to use the full slide
  width with generous spacing, instead of cramming everyone into the left
  half of the slide with illegibly small text.

## What's still PNG-only / unchanged from the previous version

- **Flexible Excel headers**, **fuzzy photo-filename matching**, **PPT-
  equivalent font sizing**, and **automatic light/dark theme detection**
  all carry over unchanged and apply to both the PPTX and PNG output paths.

## Setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

Open the local URL Streamlit prints (usually http://localhost:8501).

(Carried over from the previous version, unchanged: flexible Excel header
matching, PowerPoint-equivalent pt font sizing, and automatic light/dark
theme detection — see "What's new in this version" above for details on
each.)

## How to use

1. **Step 1 — Template & Mask** (do this once per event/theme)
   - Upload the base template PNG.
   - Upload a mask PNG: a single shape (circle, hexagon, anything) drawn in
     white/opaque on a transparent (or dark) background. This defines what
     shape every speaker photo gets cropped into. Draw only ONE shape — the
     app reuses it for every speaker slot.

2. **Step 2 — Session Details**
   - Session name, hall name, date.

3. **Step 3 — Speakers**
   - Choose **Bulk upload** (recommended for most sessions) or **Manual entry**
     (fine for 1-2 speakers).
   - **Bulk upload:**
     - Download the Excel template (or use your own sheet — headers don't
       need to match exactly, see "What's new" above), fill in speaker
       details, one row per speaker. The **Role** column is optional free
       text — write "Moderator", "Chief Guest", "Keynote Speaker", etc., or
       leave it blank for a regular speaker (no label shown on the slide).
     - Bulk-upload all speaker photos at once. Filenames don't need to match
       exactly — `ravi_singh.jpg`, `Singh_Ravi.PNG`, `Singh, Ravi.jpg` all
       match "Ravi Singh" automatically (matching ignores order, case,
       separators, and common honorifics like Dr./Shri/IRS).
     - **Review the match table** that appears — every speaker shows its
       auto-matched photo. Low-confidence matches are flagged; use the
       dropdown to fix any wrong or missing match before continuing. This
       review step exists specifically so a bad auto-match never goes
       through silently.
   - **Manual entry:** same as before — type each speaker's details and
     upload their photo individually. There's a "Role label" field per
     speaker for the same free-text role behavior as the Excel column.

4. **Step 4 — Process & Review**
   - Click "Process Photos." The app auto-detects each face, crops to a
     head-to-collar bust shot, and lightly sharpens/enhances it. Speakers
     are auto-ordered Moderator/Chair first by default.
   - **Review the thumbnails.** If a photo couldn't be auto-detected
     confidently, it's flagged with a warning — re-upload a clearer/more
     front-facing photo for that speaker and re-process if needed.
   - **Reorder if needed.** Use the ↑ / ↓ buttons under each speaker's
     photo to rearrange the lineup into your preferred order.

5. **Step 5 — Generate**
   - Choose **PowerPoint (.pptx)** (recommended — fully editable in
     PowerPoint) or **PNG** (flat image, view-only).
   - Click "Generate Slide" and download the result.

## What this version does NOT do (by design, for now)

- **No background removal** on speaker photos — backgrounds are kept as
  submitted, intentionally (per project decision — removal models are
  unreliable enough on amateur photos that the failure mode looked worse
  than the problem it solves).
- **Fixed grid layouts, not true dynamic reflow** — layouts are predefined
  per speaker count (1 through 8). This was a deliberate v1 simplification
  to get something working and visually predictable fast. Adding a new
  count = adding one entry to `core/layout_engine.py`, no other code
  changes needed. A future version can replace this with true auto-tiling
  if needed.
- **PPTX photo crops are flattened, not live-editable shapes.** PowerPoint
  can only natively crop pictures to its own built-in preset shapes (oval,
  hexagon, rectangle, etc.) — it cannot crop to an arbitrary custom
  silhouette the way the designer's mask PNG can. Per project decision, we
  preserve the EXACT designer mask shape by pre-flattening the cropped
  photo with Pillow before inserting it as a PowerPoint picture. This means
  the photo's position/size can be adjusted live in PowerPoint, and the
  photo itself can be swapped out, but the crop shape itself is baked in
  (not a live PowerPoint crop you can drag-adjust).
- **One slide per generated file currently** — multi-session decks (each
  session as its own slide in one file) are a planned near-term follow-up,
  not yet built.
- **Single mask shape only** — if a mask file contains multiple disconnected
  shapes, the app uses the largest one and warns you. It does not yet
  support masks with deliberately different shapes per slot.
- **`CANVAS_TOP_RATIO` is currently a tuned constant** (in
  `core/slide_compositor.py` and `core/pptx_compositor.py`), calibrated
  against the one sample template provided. If a new template has a taller
  or shorter header band, this will need adjusting (or it'll need to
  become a per-template setting in the UI) — that's flagged as a
  near-term follow-up.

## Project structure

```
app.py                      Streamlit UI (what staff interacts with)
core/
  mask_parser.py            Extracts the slot shape from a designer mask PNG
  photo_processor.py        Face detection, bust crop, sharpen/enhance
  layout_engine.py          Predefined grid layouts by speaker count (no
                             role/label data - that comes only from speaker
                             data, never from slot position - see bug note below)
  slide_compositor.py       PNG assembly: template + slots + photos + text,
                             PPT-equivalent pt sizing, theme-aware coloring
  pptx_compositor.py        PPTX assembly: same layout/theme logic, but
                             emits real editable PowerPoint text boxes and
                             picture shapes instead of flat pixels
  theme_detector.py         Samples template background to pick readable,
                             theme-appropriate text colors automatically
  speaker_sheet.py          Reads the speaker Excel sheet (flexible headers,
                             free-text Role column)
  name_matcher.py           Fuzzy-matches photo filenames to speaker names
assets/
  speaker_list_template.xlsx  Downloadable Excel template for bulk speaker entry
requirements.txt
```

## Bug fix note: role labels were previously tied to slot position

An earlier version hardcoded `role_label="MODERATOR"` onto whichever
speaker happened to occupy layout slot index 0 (visually, "whoever's
listed/positioned first"), completely independent of which speaker's data
actually had the moderator flag set. This caused real mislabeling (e.g. a
head-of-state shown as "Moderator" while the actual moderator further down
the list showed no label at all). This has been fixed: `Slot` objects no
longer carry any role/label data at all, and the displayed label comes
ONLY from each speaker's own `role` field. Worth specifically re-testing
this with your real session data after upgrading, given how serious a
live, on-screen mislabeling mistake like this would be.

## Known rough edges to watch for in testing

- Face detection (OpenCV Haar cascade) is reliable on clear, front-facing
  photos but can miss side profiles, sunglasses, heavy shadows, or very
  low-resolution images. This is exactly what the Step 4 review grid is
  for — it's intentionally a manual checkpoint, not a fully blind pipeline.
- **Speaker names never wrap** — by design, a very long name shrinks to
  fit its slot on one line rather than wrapping to two, since wrapping
  risked colliding with the next row in multi-row layouts. At extreme
  name lengths in narrow slots (7-8 speaker layouts), this may look
  noticeably smaller than other speakers' names in the same slide.
- **PPTX font-fit sizing uses an estimate, not exact measurement.**
  PowerPoint doesn't expose true text-width measurement without rendering,
  so `pptx_compositor.py` uses a standard average-character-width heuristic
  to decide if a name needs to shrink to fit one line. It's deliberately
  generous/safe, but isn't pixel-exact like the PNG path (which can measure
  real rendered text via Pillow). If a name looks slightly off in the
  generated pptx, it's a one-click manual font-size tweak in PowerPoint.
- **Fuzzy name matching** works well for typical cases (reordered names,
  different separators/case, common honorific suffixes like "IRS"/"Dr.")
  but isn't magic — completely generic filenames (`IMG_2024.jpg`) or two
  speakers with very similar names sharing a first or last name can produce
  an ambiguous or empty match. The Step 3 review table is the safety net:
  always glance over the assigned photo next to each name before moving on,
  since a wrong photo-to-name assignment going live on the big screen is
  the one mistake worth double-checking for.
- **Flexible Excel header matching** covers common phrasings (Name/Speaker/
  Panelist, Title/Designation/Position, Company/Organisation/Affiliation,
  Role/Moderator/Chair) but isn't infinitely smart — a sheet with genuinely
  unconventional headers (e.g. just "Info" or "Details") won't be
  recognized. If a column isn't detected, that field is simply left blank
  for all speakers rather than erroring out (except the name column, which
  is required).
- **Theme detection samples a fixed region** near the title/caption areas
  to decide light vs. dark. A template with a highly varied or textured
  background in that exact region (e.g. a busy photo background) could
  occasionally pick a less-than-ideal color. Worth a quick visual check
  on any new template style the first time it's used.
