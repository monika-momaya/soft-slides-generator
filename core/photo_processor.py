"""
photo_processor.py

Takes a raw speaker-submitted photo (any size, any aspect ratio, any
quality) and produces a clean, consistent "head-to-collar" bust crop,
sharpened and lightly enhanced - WITHOUT background removal (kept
intentionally, per requirements: backgrounds stay as submitted).

Approach:
1. Detect the face using OpenCV's built-in Haar cascade (no extra model
   downloads needed, works offline).
2. From the face box, estimate a bust crop region: roughly from just above
   the hairline down to where a collar/tie would sit (~2.1x the face
   height below the eyes, centered on the face horizontally with a bit
   of shoulder margin).
3. If no face is detected (rare, but happens with side profiles, sunglasses,
   group-photo artifacts, very low quality), fall back to a center-weighted
   crop and flag it so the review step can catch it.
4. Sharpen + mild auto-contrast/levels for consistency across photos of
   very different source quality.
"""

import cv2
import numpy as np
from PIL import Image, ImageEnhance, ImageFilter

_face_cascade = cv2.CascadeClassifier(
    cv2.data.haarcascades + "haarcascade_frontalface_default.xml"
)


class ProcessedPhoto:
    def __init__(self, image: Image.Image, face_detected: bool, note: str = ""):
        self.image = image          # final processed RGB PIL image (bust crop)
        self.face_detected = face_detected
        self.note = note            # human-readable flag for the review step


def _detect_face_box(bgr_image: np.ndarray):
    """Return (x, y, w, h) of the largest detected face, or None."""
    gray = cv2.cvtColor(bgr_image, cv2.COLOR_BGR2GRAY)
    faces = _face_cascade.detectMultiScale(
        gray, scaleFactor=1.1, minNeighbors=5, minSize=(60, 60)
    )
    if len(faces) == 0:
        return None
    # If multiple faces detected (e.g. submitted photo has other people in
    # background), assume the largest face is the intended subject.
    return max(faces, key=lambda f: f[2] * f[3])


def _bust_crop_box(face_box, img_w, img_h):
    """
    Given a face bounding box, compute a bust-crop rectangle that goes
    from just above the hairline down to roughly the collar/tie line,
    with shoulder margin on either side.

    Tuned ratios (relative to face height `fh`):
    - top margin above face top: 0.55 * fh   (hair/forehead room)
    - bottom extension below face bottom: 1.35 * fh   (down to collar/chest)
    - horizontal margin each side: 0.9 * face width   (shoulder room)
    """
    fx, fy, fw, fh = face_box

    top = fy - 0.55 * fh
    bottom = fy + fh + 1.35 * fh
    left = fx - 0.9 * fw
    right = fx + fw + 0.9 * fw

    # Clamp to image bounds
    top = max(0, top)
    left = max(0, left)
    bottom = min(img_h, bottom)
    right = min(img_w, right)

    return int(left), int(top), int(right), int(bottom)


def _center_fallback_crop(img_w, img_h):
    """
    Fallback crop when no face is detected: a centered square-ish region
    biased toward the upper-middle of the frame (where a bust shot
    usually sits), so it's a reasonable default rather than a wild guess.
    """
    side = min(img_w, img_h)
    left = (img_w - side) // 2
    top = max(0, int(img_h * 0.05))
    bottom = min(img_h, top + side)
    right = left + side
    return left, top, right, bottom


def process_photo(pil_image: Image.Image) -> ProcessedPhoto:
    """
    Main entry point. Takes any PIL image, returns a ProcessedPhoto with
    a clean bust crop, sharpened and lightly color-corrected.
    """
    rgb_image = pil_image.convert("RGB")
    img_w, img_h = rgb_image.size

    bgr = cv2.cvtColor(np.array(rgb_image), cv2.COLOR_RGB2BGR)
    face_box = _detect_face_box(bgr)

    if face_box is not None:
        left, top, right, bottom = _bust_crop_box(face_box, img_w, img_h)
        note = ""
        face_detected = True
    else:
        left, top, right, bottom = _center_fallback_crop(img_w, img_h)
        note = (
            "No face detected automatically - crop is a best-guess center "
            "crop. Please review this photo."
        )
        face_detected = False

    cropped = rgb_image.crop((left, top, right, bottom))

    # --- Enhancement pass: mild sharpen + auto-levels for consistency ---
    cropped = cropped.filter(ImageFilter.UnsharpMask(radius=2, percent=60, threshold=3))
    cropped = ImageEnhance.Contrast(cropped).enhance(1.08)
    cropped = ImageEnhance.Color(cropped).enhance(1.05)
    cropped = ImageEnhance.Brightness(cropped).enhance(1.02)

    return ProcessedPhoto(image=cropped, face_detected=face_detected, note=note)
