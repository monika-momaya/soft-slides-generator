"""
name_matcher.py

Matches bulk-uploaded photo files to speaker rows from the Excel sheet,
using fuzzy name comparison - because photo filenames won't exactly match
the "Name" column. They might be "ravi_singh.jpg", "Singh_Ravi.JPG",
"Ravi Singh.png", etc.

Approach:
1. Normalize both the speaker name and each filename into a "word set"
   (lowercase, strip extension, split on common separators, drop common
   honorifics/titles that sometimes leak into filenames).
2. Score every (speaker, photo) pair by word-set overlap (Jaccard-style).
3. Greedily assign the highest-confidence pairs first, so strong matches
   aren't blocked by ambiguous ones.
4. Anything left unmatched (or matched with low confidence) is flagged for
   manual review rather than guessed silently - this matters most on a
   live event day, where a silently wrong photo-to-name match is the
   failure mode to avoid above all else.
"""

import os
import re
from dataclasses import dataclass, field
from typing import List, Optional

HONORIFICS = {
    "dr", "mr", "mrs", "ms", "shri", "smt", "prof", "professor",
    "sir", "madam", "ir", "irs", "ips", "ias",
}

# Below this score, we don't trust an automatic match - surface it for
# manual review instead of guessing.
CONFIDENCE_THRESHOLD = 0.4


KNOWN_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".heic", ".bmp", ".tiff", ".gif"}


def _normalize_to_wordset(text: str, is_filename: bool = False) -> set:
    if is_filename:
        base, ext = os.path.splitext(text)
        # Only actually strip it if it's a real image extension - otherwise
        # a name like "Dr. Smith" would lose "Smith" to a false-positive
        # "extension" split on that period.
        if ext.lower() in KNOWN_IMAGE_EXTENSIONS:
            text = base
    text = text.lower()
    text = re.sub(r"[_\-\.]+", " ", text)
    text = re.sub(r"[^a-z0-9 ]+", " ", text)
    words = {w for w in text.split() if w and w not in HONORIFICS}
    return words


def _similarity(set_a: set, set_b: set) -> float:
    """Jaccard-style overlap, but biased to reward full containment
    (handles partial filenames like just 'ravi.jpg' matching 'ravi singh')."""
    if not set_a or not set_b:
        return 0.0
    intersection = set_a & set_b
    if not intersection:
        return 0.0
    smaller = min(len(set_a), len(set_b))
    # Reward how much of the SMALLER set is covered - this is what lets a
    # short filename like "ravi.jpg" still score well against "ravi singh"
    containment = len(intersection) / smaller
    jaccard = len(intersection) / len(set_a | set_b)
    return 0.6 * containment + 0.4 * jaccard


@dataclass
class MatchResult:
    speaker_name: str
    matched_filename: Optional[str] = None
    confidence: float = 0.0
    needs_review: bool = True
    candidates: List[str] = field(default_factory=list)  # other plausible filenames, for a manual dropdown


def match_photos_to_speakers(speaker_names: List[str], photo_filenames: List[str]) -> List[MatchResult]:
    """
    speaker_names: list of names as entered in the Excel sheet, in order.
    photo_filenames: list of filenames from the bulk photo upload.

    Returns one MatchResult per speaker, in the same order as speaker_names.
    """
    speaker_sets = [_normalize_to_wordset(name, is_filename=False) for name in speaker_names]
    photo_sets = [_normalize_to_wordset(fname, is_filename=True) for fname in photo_filenames]

    # Build full score matrix
    scores = []  # list of (score, speaker_idx, photo_idx)
    for s_idx, s_set in enumerate(speaker_sets):
        for p_idx, p_set in enumerate(photo_sets):
            score = _similarity(s_set, p_set)
            if score > 0:
                scores.append((score, s_idx, p_idx))

    scores.sort(key=lambda x: -x[0])

    assigned_speaker = {}
    used_photos = set()

    for score, s_idx, p_idx in scores:
        if s_idx in assigned_speaker or p_idx in used_photos:
            continue
        assigned_speaker[s_idx] = (photo_filenames[p_idx], score)
        used_photos.add(p_idx)

    results = []
    for s_idx, name in enumerate(speaker_names):
        if s_idx in assigned_speaker:
            fname, score = assigned_speaker[s_idx]
            results.append(MatchResult(
                speaker_name=name,
                matched_filename=fname,
                confidence=round(score, 2),
                needs_review=score < CONFIDENCE_THRESHOLD,
            ))
        else:
            # No match at all - offer the unused photos as manual candidates
            unused = [photo_filenames[i] for i in range(len(photo_filenames)) if i not in used_photos]
            results.append(MatchResult(
                speaker_name=name,
                matched_filename=None,
                confidence=0.0,
                needs_review=True,
                candidates=unused,
            ))
    return results
