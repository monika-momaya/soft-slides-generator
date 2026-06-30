"""
theme_detector.py

Auto-detects text colors based on the actual template background, instead
of hardcoding white text. Since every event's template is different (dark
gradients, light backgrounds, photos, etc.), we sample the real pixels
behind where text will be drawn and pick readable colors from that.

Approach:
- Sample a region of the template (e.g. where the session title will sit).
- Compute average luminance to classify it as "dark" or "light".
- Return a small palette: primary text color (high contrast, e.g. white
  on dark / near-black on light) and an accent color (for speaker names -
  vibrant, theme-appropriate, but still readable) and a muted/secondary
  color (for title+company lines).

This is intentionally simple color theory (luminance threshold + fixed
accent palettes per light/dark), not full template-driven palette
extraction - that's a reasonable place to get fancier later if needed,
but this covers "looks right on light vs. dark templates automatically"
without requiring any manual per-template flag.
"""

from dataclasses import dataclass
from PIL import Image
import numpy as np


@dataclass
class TextTheme:
    is_dark_background: bool
    primary_color: tuple        # for session title - strong contrast
    secondary_color: tuple      # for hall/date meta line - softer contrast
    name_accent_color: tuple    # for speaker names - vibrant, bold
    caption_color: tuple        # for title+company under each speaker - soft, no strong color
    role_label_color: tuple     # for "MODERATOR" style labels


# Accent palettes tuned to read well against each background type.
# (R, G, B) - alpha is added by callers as needed.
_DARK_BG_PALETTE = dict(
    primary_color=(255, 255, 255),
    secondary_color=(220, 220, 230),
    name_accent_color=(255, 200, 60),     # warm gold/amber - common in event branding, reads well on dark
    caption_color=(210, 210, 220),
    role_label_color=(255, 210, 60),
)

_LIGHT_BG_PALETTE = dict(
    primary_color=(20, 20, 30),
    secondary_color=(60, 60, 70),
    name_accent_color=(180, 30, 60),      # deep vibrant red/maroon - reads well on light, still feels "branded"
    caption_color=(70, 70, 80),
    role_label_color=(190, 40, 70),
)


def _region_luminance(image: Image.Image, box: tuple) -> float:
    """Average perceived luminance (0=black, 255=white) of a region."""
    region = image.convert("RGB").crop(box)
    arr = np.array(region).astype(np.float32)
    # Standard luminance weighting
    luminance = 0.2126 * arr[:, :, 0] + 0.7152 * arr[:, :, 1] + 0.0722 * arr[:, :, 2]
    return float(luminance.mean())


def detect_theme(template: Image.Image, sample_box: tuple, dark_threshold: float = 130.0) -> TextTheme:
    """
    template: the full template image.
    sample_box: (left, top, right, bottom) region to sample - typically the
                area where the session title / speaker captions will render.
    dark_threshold: luminance below this = classified as dark background.
    """
    luminance = _region_luminance(template, sample_box)
    is_dark = luminance < dark_threshold
    palette = _DARK_BG_PALETTE if is_dark else _LIGHT_BG_PALETTE
    return TextTheme(is_dark_background=is_dark, **palette)
