"""
layout_engine.py

v1 approach (deliberately simple, per project decision): rather than true
dynamic reflow, we use a small set of PREDEFINED grid layouts keyed by
speaker count. Each layout defines normalized (0-1) slot positions within
the open canvas area, so it scales to any canvas size.

IMPORTANT: slots carry NO role/label information. Early versions hardcoded
role_label="MODERATOR" on slot index 0, assuming whoever was listed first
would always be the moderator - this was a real bug (a speaker's displayed
role must come ONLY from their own data, never from their position in the
list). Role/label text is entirely the caller's responsibility now, driven
by each speaker's own role field.

Per a later round of feedback, layouts are redesigned to:
- Always use the FULL slide width (earlier 7/8-speaker layouts only used
  ~4 columns' worth of width, leaving a large empty gap on the right -
  this was a real bug, not a style choice).
- Pair with sort_speakers_moderator_first() below, so moderator/chair-type
  speakers visually appear first/leftmost on the slide, matching the
  reference slide convention - this is a SORTING concern, not a layout
  concern, so it lives in a separate function the caller applies BEFORE
  calling get_layout(), keeping slot position fully decoupled from any
  notion of role (preserving the earlier bug fix's spirit: slots still
  carry no role data).

Adding a new layout = adding one entry to LAYOUTS. No other code changes
needed elsewhere in the app.

NOTE for future upgrade path: if/when true auto-tile reflow is wanted,
this module is the ONLY place that needs to change - everything downstream
(photo processing, compositing) just consumes whatever slot list this
module returns, regardless of how it was computed.
"""

from dataclasses import dataclass
from typing import List, Dict


@dataclass
class Slot:
    """A single photo+caption slot, in NORMALIZED coordinates (0.0-1.0)
    relative to the open canvas area (the space below the template header)."""
    x: float       # left edge
    y: float        # top edge
    w: float        # width
    h: float        # height (photo only; caption text goes below this)


# Each layout is a list of Slot objects for a specific total speaker count.
# Coordinates designed for a roughly 16:9 open canvas area, and always
# spread across the FULL width (x + w reaches ~0.98 for the rightmost slot)
# rather than clustering in part of the slide.
LAYOUTS = {
    1: [
        Slot(x=0.36, y=0.10, w=0.28, h=0.55),
    ],
    2: [
        Slot(x=0.14, y=0.12, w=0.32, h=0.50),
        Slot(x=0.54, y=0.12, w=0.32, h=0.50),
    ],
    3: [
        Slot(x=0.04, y=0.12, w=0.29, h=0.48),
        Slot(x=0.355, y=0.12, w=0.29, h=0.48),
        Slot(x=0.67, y=0.12, w=0.29, h=0.48),
    ],
    4: [
        Slot(x=0.02, y=0.12, w=0.23, h=0.46),
        Slot(x=0.27, y=0.12, w=0.23, h=0.46),
        Slot(x=0.52, y=0.12, w=0.23, h=0.46),
        Slot(x=0.77, y=0.12, w=0.21, h=0.46),
    ],
    5: [
        Slot(x=0.01, y=0.04, w=0.19, h=0.40),
        Slot(x=0.21, y=0.04, w=0.19, h=0.40),
        Slot(x=0.41, y=0.04, w=0.19, h=0.40),
        Slot(x=0.61, y=0.04, w=0.19, h=0.40),
        Slot(x=0.81, y=0.04, w=0.18, h=0.40),
    ],
    6: [
        Slot(x=0.01, y=0.02, w=0.193, h=0.42),
        Slot(x=0.205, y=0.02, w=0.193, h=0.42),
        Slot(x=0.40, y=0.02, w=0.193, h=0.42),
        Slot(x=0.595, y=0.02, w=0.193, h=0.42),
        Slot(x=0.01, y=0.52, w=0.193, h=0.42),
        Slot(x=0.205, y=0.52, w=0.193, h=0.42),
    ],
    7: [
        # 4 across top row, 3 across bottom row, full slide width both rows
        Slot(x=0.01, y=0.02, w=0.23, h=0.42),
        Slot(x=0.255, y=0.02, w=0.23, h=0.42),
        Slot(x=0.50, y=0.02, w=0.23, h=0.42),
        Slot(x=0.745, y=0.02, w=0.23, h=0.42),
        Slot(x=0.135, y=0.52, w=0.23, h=0.42),
        Slot(x=0.38, y=0.52, w=0.23, h=0.42),
        Slot(x=0.625, y=0.52, w=0.23, h=0.42),
    ],
    8: [
        # 4 across top row, 4 across bottom row, full slide width both rows
        Slot(x=0.01, y=0.02, w=0.23, h=0.42),
        Slot(x=0.255, y=0.02, w=0.23, h=0.42),
        Slot(x=0.50, y=0.02, w=0.23, h=0.42),
        Slot(x=0.745, y=0.02, w=0.23, h=0.42),
        Slot(x=0.01, y=0.52, w=0.23, h=0.42),
        Slot(x=0.255, y=0.52, w=0.23, h=0.42),
        Slot(x=0.50, y=0.52, w=0.23, h=0.42),
        Slot(x=0.745, y=0.52, w=0.23, h=0.42),
    ],
}

MAX_SUPPORTED = max(LAYOUTS.keys())


def get_layout(speaker_count: int) -> List[Slot]:
    """
    Return the slot list for the given speaker count. If the count exceeds
    what we have a hand-tuned layout for, we raise a clear error rather
    than silently producing an ugly result - the calling UI should surface
    this so staff knows to ask for a layout to be added, or split the
    panel across two slides.
    """
    if speaker_count < 1:
        raise ValueError("Need at least 1 speaker to generate a layout.")
    if speaker_count not in LAYOUTS:
        raise ValueError(
            f"No predefined layout for {speaker_count} speakers yet. "
            f"Currently supported: 1-{MAX_SUPPORTED} speakers. "
            "Consider splitting this panel across two slides, or ask "
            "to have a new layout added for this count."
        )
    return LAYOUTS[speaker_count]


# Role values (lowercased) treated as "should appear first/leftmost" when
# auto-sorting. This is intentionally a short, specific list rather than
# "anything non-default" - a Chief Guest or Keynote Speaker is notable but
# isn't necessarily meant to visually lead the lineup the way a Moderator/
# Chair conventionally does in panel-style sessions (matches the reference
# slides, where the moderator specifically anchors the leftmost position).
MODERATOR_PRIORITY_ROLES = {"moderator", "chair", "co-moderator", "co-chair"}


def sort_speakers_moderator_first(speakers: List[Dict]) -> List[Dict]:
    """
    Returns a NEW list with moderator/chair-role speakers moved to the
    front, preserving relative order within each group otherwise (stable
    sort) - so among multiple moderators, or among regular speakers, the
    original (e.g. user-set/reordered) order is kept.

    Each speaker dict is expected to have a "role" key (free text, as
    produced by speaker_sheet.read_speaker_sheet or the manual entry form).
    """
    def is_priority(sp):
        return sp.get("role", "").strip().lower() in MODERATOR_PRIORITY_ROLES

    priority = [sp for sp in speakers if is_priority(sp)]
    rest = [sp for sp in speakers if not is_priority(sp)]
    return priority + rest
