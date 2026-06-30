"""
slide_compositor.py

Final assembly step: takes the base template, the extracted slot shape,
the layout (slot positions), processed speaker photos, and text fields
(session name, hall, date, speaker name/title/company) and renders the
finished, ready-to-project slide.

Font sizing convention:
- Sizes are specified in "pt", matching how these slides are normally
  authored in PowerPoint at a 1280x720 (96 DPI) reference canvas - i.e.
  the same assumption PowerPoint/web use for on-screen pt sizing, NOT a
  print-points conversion. 1pt = 1.3333px at that reference resolution.
- The actual template image may be a different resolution (e.g. 1616x910),
  so pt sizes are scaled by (template_width / 1280) to stay visually
  consistent with how they'd look authored in PowerPoint, regardless of
  the exported template's pixel dimensions.

Text styling rules (per project spec):
- Session name: bold, 20pt ideal, never auto-shrunk below 16pt. If still
  too long at 16pt, it wraps to a second line rather than shrinking further
  or being cut off - it must never become illegibly small.
- Speaker name: bold, vibrant/theme-aware accent color, sized 4pt LARGER
  than the title/company line (not just "less small").
- Title + Company: regular weight (not bold), no special color - follows
  the same light/dark text logic as the rest of the template's theme,
  sized 3pt smaller than the speaker name.
- All template-text coloring (session name, captions) is auto-detected
  per template: a region's average brightness decides whether to use
  light or dark, theme-appropriate text colors - no manual flag needed,
  since templates change per event.
"""

from dataclasses import dataclass
from typing import List, Optional
from PIL import Image, ImageDraw, ImageFont

from core.layout_engine import Slot
from core.mask_parser import SlotShape
from core.theme_detector import detect_theme

# Fraction of the template's height taken up by the fixed branded header
# (logo/date/banner band). Below this is the open canvas for speakers.
# NOTE: tuned against the one sample template provided; ideally this
# becomes auto-detected or per-template configurable in a future pass.
CANVAS_TOP_RATIO = 0.32

CAPTION_GAP_PX_RATIO = 0.012      # gap between photo bottom and name text, as a fraction of template height
NAME_TO_TITLE_GAP_RATIO = 0.004
TITLE_TO_COMPANY_GAP_RATIO = 0.0

# --- PPT-equivalent pt sizing rules ---
REFERENCE_WIDTH_PX = 1280   # PowerPoint 16:9 standard slide width at 96 DPI
PT_TO_PX_AT_REFERENCE = 1.3333  # 96 DPI: 1pt = 1.3333px, same convention as PowerPoint/web

SESSION_NAME_IDEAL_PT = 20
SESSION_NAME_MIN_PT = 16
META_LINE_PT = 13

SPEAKER_NAME_PT = 18        # "4pt larger than title/company" -> title/company = 14pt
TITLE_COMPANY_PT = 14
ROLE_LABEL_PT = 13

# Role values that are treated as "no special role" and therefore NOT shown
# as a label above the speaker's photo - matches the convention in the
# reference slides, where only notable roles (Moderator, Chief Guest, etc.)
# get a label and ordinary panelists/speakers don't.
DEFAULT_LABEL_VALUES_TO_HIDE = {"speaker", "panelist", ""}


@dataclass
class SpeakerInfo:
    name: str
    title: str = ""
    company: str = ""
    photo: Optional[Image.Image] = None   # already-processed bust crop (RGB)
    role_label: str = ""                  # free text, e.g. "Moderator", "Chief Guest",
                                           # "Keynote Speaker". Default-ish values like
                                           # "Speaker"/"Panelist" are NOT shown as a label
                                           # above the photo (matches reference slide
                                           # convention - only notable roles get a label).
                                           # See DEFAULT_LABEL_VALUES_TO_HIDE below.


def _pt_to_px(pt: float, template_width: int) -> int:
    """
    Convert a PowerPoint-style pt size to pixels, scaled to the template's
    actual resolution so it looks the same as it would authored at the
    standard 1280px-wide 16:9 PowerPoint reference canvas.
    """
    scale = template_width / REFERENCE_WIDTH_PX
    return max(1, round(pt * PT_TO_PX_AT_REFERENCE * scale))


def _font(size_px: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    """
    Load a font at the given pixel size. Falls back through a few common
    paths so this works across different environments without bundling
    font files. Production deployments should bundle a specific
    brand-approved font.
    """
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold
        else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf" if bold
        else "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, size_px)
        except (OSError, IOError):
            continue
    return ImageFont.load_default()


def _text_width(draw, text, font) -> int:
    bbox = draw.textbbox((0, 0), text, font=font)
    return bbox[2] - bbox[0]


def _wrap_to_fit(draw, text, font, max_width: int) -> List[str]:
    """Greedy word-wrap text into lines that each fit within max_width."""
    words = text.split()
    if not words:
        return [""]
    lines = []
    current = words[0]
    for word in words[1:]:
        candidate = f"{current} {word}"
        if _text_width(draw, candidate, font) <= max_width:
            current = candidate
        else:
            lines.append(current)
            current = word
    lines.append(current)
    return lines


def _fit_text_font(draw: ImageDraw.ImageDraw, text: str, max_width: int,
                    start_size: int, min_size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    """Shrink font size until `text` fits within `max_width` pixels, or hit min_size."""
    size = start_size
    while size > min_size:
        font = _font(size, bold=bold)
        if _text_width(draw, text, font) <= max_width:
            return font
        size -= 1
    return _font(min_size, bold=bold)


def _draw_centered_text(draw, text, center_x, top_y, font, color, max_width=None,
                         min_size_px=None, bold=False):
    """
    Draw text horizontally centered at center_x on a SINGLE line, returns
    the height used. If max_width is given and the text overflows it, the
    font is shrunk down (never wrapped) until it fits, or until min_size_px
    is hit - whichever comes first. Single-line-only is a deliberate choice:
    wrapping speaker captions in fixed-grid layouts risks colliding with the
    next row, so a slightly smaller name is preferred over an unpredictable
    multi-line block.
    """
    if not text:
        return 0

    if max_width is not None and _text_width(draw, text, font) > max_width:
        size = font.size
        floor = min_size_px or max(10, int(size * 0.6))
        while size > floor:
            size -= 1
            font = _font(size, bold=bold)
            if _text_width(draw, text, font) <= max_width:
                break

    bbox = draw.textbbox((0, 0), text, font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]
    draw.text((center_x - text_w / 2, top_y), text, font=font, fill=color)
    return text_h


def _apply_mask_shape(photo: Image.Image, slot_shape: SlotShape, available_w: int, available_h: int) -> Image.Image:
    """
    Fits the slot shape (preserving ITS OWN aspect ratio - e.g. a circle
    must stay a circle, not become an oval) within the available
    (available_w, available_h) box, centers it, then crops/resizes the
    speaker photo to "cover" that shape (like CSS object-fit: cover)
    before applying the mask.

    Returns an RGBA image sized (available_w, available_h) with the
    masked photo centered and transparent padding around it - safe to
    paste directly at the slot's top-left corner.
    """
    # Fit the shape's native aspect ratio within the available box
    shape_ratio = slot_shape.aspect_ratio
    box_ratio = available_w / available_h

    if shape_ratio > box_ratio:
        # shape is relatively wider than the box -> width is the binding constraint
        shape_w = available_w
        shape_h = int(round(shape_w / shape_ratio))
    else:
        shape_h = available_h
        shape_w = int(round(shape_h * shape_ratio))

    offset_x = (available_w - shape_w) // 2
    offset_y = (available_h - shape_h) // 2

    # Resize photo to "cover" the shape's box (crop to fill, preserving photo aspect)
    photo_ratio = photo.width / photo.height
    target_ratio = shape_w / shape_h

    if photo_ratio > target_ratio:
        new_h = shape_h
        new_w = int(new_h * photo_ratio)
    else:
        new_w = shape_w
        new_h = int(new_w / photo_ratio)

    resized = photo.resize((new_w, new_h), Image.LANCZOS)
    crop_left = (new_w - shape_w) // 2
    crop_top = (new_h - shape_h) // 2
    cropped = resized.crop((crop_left, crop_top, crop_left + shape_w, crop_top + shape_h))

    mask_resized = slot_shape.resized(shape_w, shape_h)

    masked_shape = Image.new("RGBA", (shape_w, shape_h), (0, 0, 0, 0))
    masked_shape.paste(cropped.convert("RGBA"), (0, 0))
    masked_shape.putalpha(mask_resized)

    result = Image.new("RGBA", (available_w, available_h), (0, 0, 0, 0))
    result.paste(masked_shape, (offset_x, offset_y), masked_shape)
    return result


def compose_slide(
    template: Image.Image,
    slot_shape: SlotShape,
    slots: List[Slot],
    speakers: List[SpeakerInfo],
    session_name: str = "",
    hall_name: str = "",
    date_str: str = "",
) -> Image.Image:
    """
    Render the final slide. `slots` and `speakers` must be the same length
    (caller's responsibility - typically slots come from layout_engine.get_layout(len(speakers))).
    """
    if len(slots) != len(speakers):
        raise ValueError(
            f"Slot count ({len(slots)}) does not match speaker count ({len(speakers)})."
        )

    canvas = template.convert("RGBA").copy()
    W, H = canvas.size
    draw = ImageDraw.Draw(canvas)

    canvas_top = int(H * CANVAS_TOP_RATIO)

    # Sample the template's own colors near the title area and the speaker
    # caption area separately - a template could plausibly have a light
    # header and a darker open canvas (or vice versa), so we detect each
    # region independently rather than assuming one theme for the whole slide.
    title_theme = detect_theme(template, (0, canvas_top, W, min(H, canvas_top + int(H * 0.12))))
    caption_theme = detect_theme(template, (0, canvas_top + int(H * 0.15), W, H))

    caption_gap_px = int(H * CAPTION_GAP_PX_RATIO)
    name_to_title_gap_px = int(H * NAME_TO_TITLE_GAP_RATIO)
    title_to_company_gap_px = int(H * TITLE_TO_COMPANY_GAP_RATIO)

    # --- Session title block (top-left of open canvas) ---
    # Bold, 20pt ideal, never auto-shrunk below 16pt - wraps to a second
    # line instead of shrinking further, so it never becomes illegible.
    title_reserved_height = 0
    if session_name:
        max_title_width = int(W * 0.94)
        ideal_px = _pt_to_px(SESSION_NAME_IDEAL_PT, W)
        min_px = _pt_to_px(SESSION_NAME_MIN_PT, W)

        title_font = _fit_text_font(draw, session_name, max_title_width, ideal_px, min_px, bold=True)
        # If even at the minimum size it doesn't fit on one line, wrap it
        # rather than shrinking past the 16pt floor.
        if _text_width(draw, session_name, title_font) > max_title_width:
            title_font = _font(min_px, bold=True)
            lines = _wrap_to_fit(draw, session_name, title_font, max_title_width)
        else:
            lines = [session_name]

        line_y = canvas_top + int(H * 0.005)
        line_height = 0
        for line in lines[:2]:  # cap at 2 lines to keep layout predictable
            bbox = draw.textbbox((0, 0), line, font=title_font)
            line_height = bbox[3] - bbox[1]
            draw.text((int(W * 0.03), line_y), line, font=title_font, fill=title_theme.primary_color)
            line_y += int(line_height * 1.25)

        title_reserved_height = (line_y - canvas_top) + int(H * 0.01)

    meta_parts = [p for p in [hall_name, date_str] if p]
    if meta_parts:
        meta_text = "   |   ".join(meta_parts)
        meta_font = _font(_pt_to_px(META_LINE_PT, W))
        draw.text((int(W * 0.03), canvas_top + title_reserved_height), meta_text,
                   font=meta_font, fill=title_theme.secondary_color)
        meta_bbox = draw.textbbox((0, 0), meta_text, font=meta_font)
        title_reserved_height += (meta_bbox[3] - meta_bbox[1]) + int(H * 0.015)

    speaker_area_top = canvas_top + title_reserved_height
    canvas_height = H - speaker_area_top

    # --- Speaker slots ---
    name_font = _font(_pt_to_px(SPEAKER_NAME_PT, W), bold=True)
    caption_font = _font(_pt_to_px(TITLE_COMPANY_PT, W), bold=False)
    role_font = _font(_pt_to_px(ROLE_LABEL_PT, W), bold=True)

    for slot, speaker in zip(slots, speakers):
        slot_x = int(slot.x * W)
        slot_y = speaker_area_top + int(slot.y * canvas_height)
        slot_w = int(slot.w * W)
        slot_h = int(slot.h * canvas_height)

        role_text = speaker.role_label if speaker.role_label.strip().lower() not in DEFAULT_LABEL_VALUES_TO_HIDE else ""

        # Reserve a portion of the slot's allotted height for the photo and
        # leave the rest for the caption stack (name + title + company),
        # so captions never spill past this slot's space into the next row.
        role_extra_h = 0
        role_bbox = None
        if role_text:
            role_bbox = draw.textbbox((0, 0), role_text, font=role_font)
            role_extra_h = (role_bbox[3] - role_bbox[1]) + int(H * 0.012)

        caption_reserve_h = int(slot_h * 0.34)  # name + title + company + gaps
        photo_h = slot_h - role_extra_h - caption_reserve_h
        photo_top = slot_y + role_extra_h

        if role_text:
            draw.text(
                (slot_x + slot_w / 2 - (role_bbox[2] - role_bbox[0]) / 2, slot_y),
                role_text, font=role_font, fill=caption_theme.role_label_color,
            )

        if speaker.photo is not None:
            masked = _apply_mask_shape(speaker.photo, slot_shape, slot_w, int(photo_h))
            canvas.paste(masked, (slot_x, int(photo_top)), masked)

        caption_y = photo_top + photo_h + caption_gap_px
        center_x = slot_x + slot_w / 2
        caption_max_width = int(slot_w * 0.96)

        # Speaker name: bold, vibrant theme-aware accent color, larger than caption text
        used_h = _draw_centered_text(
            draw, speaker.name, center_x, caption_y, name_font, caption_theme.name_accent_color,
            max_width=caption_max_width, bold=True,
        )
        caption_y += used_h + name_to_title_gap_px

        # Title + Company: regular weight, plain theme-aware (light/dark) color, smaller than name
        used_h = _draw_centered_text(
            draw, speaker.title, center_x, caption_y, caption_font, caption_theme.caption_color,
            max_width=caption_max_width, bold=False,
        )
        caption_y += used_h + title_to_company_gap_px

        _draw_centered_text(
            draw, speaker.company, center_x, caption_y, caption_font, caption_theme.caption_color,
            max_width=caption_max_width, bold=False,
        )

    return canvas
