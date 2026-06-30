"""
mask_parser.py

Reads a designer-provided mask PNG and extracts the "slot shape" - the
reusable cutout shape used to place each speaker photo (circle, hexagon,
blob, anything). The designer draws ONE shape on a transparent background;
opaque/white pixels = where the photo should show through.

We extract:
- a tight-cropped alpha mask of just that one shape
- its native aspect ratio (so we don't distort it when resizing)
"""

from PIL import Image
import numpy as np


class SlotShape:
    """Represents one reusable photo-slot shape extracted from a mask file."""

    def __init__(self, alpha_mask: Image.Image):
        """
        alpha_mask: a single-channel (L mode) PIL Image, tightly cropped,
                    where 255 = photo visible, 0 = hidden.
        """
        self.alpha_mask = alpha_mask
        self.width, self.height = alpha_mask.size
        self.aspect_ratio = self.width / self.height

    def resized(self, target_w: int, target_h: int) -> Image.Image:
        """Return the alpha mask resized to the given dimensions."""
        return self.alpha_mask.resize((target_w, target_h), Image.LANCZOS)

    def resized_preserving_aspect(self, target_w: int) -> Image.Image:
        """Resize to a target width, preserving the original aspect ratio."""
        target_h = int(round(target_w / self.aspect_ratio))
        return self.resized(target_w, target_h)


def load_slot_shape(mask_path: str) -> SlotShape:
    """
    Load a designer mask PNG and extract the slot shape.

    Assumes the mask file contains exactly ONE shape (one connected region
    of opaque/light pixels) on an otherwise transparent or black background.
    If multiple disconnected shapes are found, the LARGEST is used (with a
    warning surfaced to caller via the returned object's `warning` attr).
    """
    img = Image.open(mask_path).convert("RGBA")
    arr = np.array(img)

    # Build a binary "shape present" map.
    # Prefer alpha channel if it has real variation; otherwise fall back
    # to brightness (handles masks that are flat white shapes on black,
    # with no transparency at all).
    alpha_channel = arr[:, :, 3]
    if alpha_channel.max() > 0 and alpha_channel.min() < alpha_channel.max():
        shape_map = alpha_channel
    else:
        # No usable alpha variation -> use luminance instead
        rgb = arr[:, :, :3].astype(np.float32)
        luminance = rgb.mean(axis=2)
        shape_map = luminance.astype(np.uint8)

    binary = (shape_map > 30).astype(np.uint8) * 255

    warning = None

    try:
        import cv2

        contours, _ = cv2.findContours(
            binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
        )
        if not contours:
            raise ValueError(
                "No shape detected in mask file. Make sure the slot shape "
                "is drawn in white/opaque pixels on a transparent or dark background."
            )

        if len(contours) > 1:
            warning = (
                f"Found {len(contours)} separate shapes in the mask file. "
                "Using the largest one as the slot shape. If your mask "
                "intentionally has multiple slots, this tool currently "
                "expects only ONE reusable shape (it auto-tiles it for "
                "multiple speakers)."
            )

        largest = max(contours, key=cv2.contourArea)
        x, y, w, h = cv2.boundingRect(largest)

    except ImportError:
        # Fallback without cv2: just use the bounding box of all opaque pixels
        ys, xs = np.where(binary > 0)
        if len(xs) == 0:
            raise ValueError(
                "No shape detected in mask file. Make sure the slot shape "
                "is drawn in white/opaque pixels on a transparent or dark background."
            )
        x, y, w, h = xs.min(), ys.min(), xs.max() - xs.min(), ys.max() - ys.min()

    # Crop tightly to the detected shape's bounding box
    cropped_alpha = Image.fromarray(binary[y:y + h, x:x + w], mode="L")

    slot = SlotShape(cropped_alpha)
    slot.warning = warning
    return slot
