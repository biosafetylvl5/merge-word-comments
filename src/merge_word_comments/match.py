"""Fuzzy match anchor text from source documents to target paragraphs."""

from typing import Optional

from rapidfuzz import fuzz

from merge_word_comments.types import Comment, MatchResult


MIN_CONTEXT_LENGTH = 80
CONTEXT_WORD_COUNT = 8


def expand_anchor_context(anchor: str, paragraph_text: str) -> str:
    """Expand short anchor text with surrounding words from the paragraph.

    For short anchors like "the" or "dog", include surrounding words
    to create a more matchable context string.
    """
    if not anchor or not paragraph_text:
        return anchor

    if len(anchor) >= MIN_CONTEXT_LENGTH:
        return anchor

    idx = paragraph_text.find(anchor)
    if idx == -1:
        return anchor

    # Expand to a character window around the anchor, keeping the original
    # text intact (no word-boundary splitting) so mid-word anchors don't
    # introduce artificial spaces.
    half_extra = (MIN_CONTEXT_LENGTH - len(anchor)) // 2
    start = max(0, idx - half_extra)
    end = min(len(paragraph_text), idx + len(anchor) + half_extra)

    # If one side hit the boundary, give the slack to the other side
    needed = MIN_CONTEXT_LENGTH - (end - start)
    if needed > 0:
        if start == 0:
            end = min(len(paragraph_text), end + needed)
        else:
            start = max(0, start - needed)

    return paragraph_text[start:end]


def _search_paragraphs(
    anchor_text: str,
    paragraphs: list[str],
    threshold: int,
    start: int = 0,
    end: Optional[int] = None,
    adaptive: bool = True,
    source_paragraph_index: Optional[int] = None,
) -> Optional[MatchResult]:
    """Search a range of paragraphs for the best match.

    Returns None if no non-empty paragraphs exist in the range.
    """
    if end is None:
        end = len(paragraphs)

    best_idx = start
    best_score = 0.0
    candidates: list[tuple[int, float]] = []
    use_token_fallback = len(anchor_text) > 200
    found_any = False

    for i in range(start, min(end, len(paragraphs))):
        para = paragraphs[i]
        if not para.strip():
            continue
        found_any = True
        score = fuzz.partial_ratio(anchor_text, para)
        if use_token_fallback and score < threshold:
            token_score = fuzz.token_sort_ratio(anchor_text, para)
            score = max(score, token_score)
        if score > best_score:
            best_score = score
            best_idx = i
        candidates.append((i, score))

    if not found_any:
        return None

    # Proximity tiebreak: among candidates within 3 points of the best,
    # prefer the one closest to source_paragraph_index.
    if source_paragraph_index is not None and candidates:
        close_candidates = [
            (idx, sc) for idx, sc in candidates
            if sc >= best_score - 3
        ]
        if len(close_candidates) > 1:
            close_candidates.sort(
                key=lambda c: abs(c[0] - source_paragraph_index)
            )
            best_idx, best_score = close_candidates[0]

    below = best_score < threshold

    # Adaptive threshold
    if below and adaptive and best_score >= threshold - 10 and candidates:
        sorted_scores = sorted((sc for _, sc in candidates), reverse=True)
        second_best = sorted_scores[1] if len(sorted_scores) > 1 else 0.0
        if best_score - second_best >= 5:
            below = False

    return MatchResult(
        comment=None,  # type: ignore[arg-type]
        target_paragraph_index=best_idx,
        score=best_score,
        anchor_offset=find_anchor_offset(anchor_text, paragraphs[best_idx]),
        below_threshold=below,
    )


def find_best_paragraph_match(
    anchor_text: str,
    paragraphs: list[str],
    threshold: int = 0,
    source_paragraph_index: Optional[int] = None,
    adaptive: bool = True,
    source_heading: str = "",
    target_heading_sections: Optional[list[tuple[str, int, int]]] = None,
) -> Optional[MatchResult]:
    """Find the paragraph that best matches the given anchor text.

    Uses rapidfuzz partial_ratio for substring matching, which handles
    cases where the anchor is a fragment of a larger paragraph.

    When *source_heading* and *target_heading_sections* are provided,
    searches within the matching heading section first. Falls back to
    full-document search if no good match is found in the scoped section.

    When *adaptive* is True, a match scoring between threshold-10 and
    threshold is accepted if it is clearly the best (margin >= 5 over
    second-best).
    """
    if not paragraphs:
        return None

    if not anchor_text:
        # Empty anchor: use source index if available, else first paragraph
        target_idx = 0
        if (
            source_paragraph_index is not None
            and 0 <= source_paragraph_index < len(paragraphs)
        ):
            target_idx = source_paragraph_index
        return MatchResult(
            comment=None,  # type: ignore[arg-type]
            target_paragraph_index=target_idx,
            score=0,
            anchor_offset=None,
            below_threshold=0 < threshold,
        )

    # Heading-scoped search: if we know the source heading, search only
    # within the matching target heading section first.
    if source_heading and target_heading_sections:
        source_kw = source_heading.upper()
        for kw, sec_start, sec_end in target_heading_sections:
            if kw.upper() == source_kw:
                scoped = _search_paragraphs(
                    anchor_text, paragraphs, threshold,
                    start=sec_start, end=sec_end,
                    adaptive=adaptive,
                    source_paragraph_index=source_paragraph_index,
                )
                if scoped is not None and not scoped.below_threshold:
                    return scoped
                break  # heading found but match was poor — fall through

    # Full-document search
    return _search_paragraphs(
        anchor_text, paragraphs, threshold,
        adaptive=adaptive,
        source_paragraph_index=source_paragraph_index,
    )


def find_anchor_offset(anchor: str, paragraph: str) -> Optional[int]:
    """Find the character offset of anchor text within a paragraph.

    First tries exact match, then falls back to fuzzy substring alignment.
    """
    if not anchor or not paragraph:
        return None

    # Try exact match first
    idx = paragraph.find(anchor)
    if idx != -1:
        return idx

    # Fuzzy sliding window: try substrings of paragraph with same length as anchor
    if len(anchor) > len(paragraph):
        return 0

    best_offset = 0
    best_score = 0.0
    window = len(anchor)

    for i in range(len(paragraph) - window + 1):
        candidate = paragraph[i:i + window]
        score = fuzz.ratio(anchor, candidate)
        if score > best_score:
            best_score = score
            best_offset = i

    if best_score < 50:
        return None

    return best_offset


def match_comments_to_target(
    comments: list[Comment],
    target_paragraphs: list[str],
    threshold: int = 80,
    adaptive: bool = True,
    target_heading_sections: Optional[list[tuple[str, int, int]]] = None,
) -> list[MatchResult]:
    """Match a list of comments to target document paragraphs.

    Each comment is matched using its anchor_context (expanded anchor text)
    to find the best target paragraph. Results include match scores
    and whether each match is below the threshold.
    """
    results = []

    for comment in comments:
        search_text = comment.anchor_context if comment.anchor_context else comment.anchor_text

        match = find_best_paragraph_match(
            search_text, target_paragraphs, threshold,
            source_paragraph_index=comment.start_paragraph_index,
            adaptive=adaptive,
            source_heading=comment.source_heading,
            target_heading_sections=target_heading_sections,
        )

        if match is None:
            results.append(MatchResult(
                comment=comment,
                target_paragraph_index=0,
                score=0,
                anchor_offset=None,
                below_threshold=True,
            ))
            continue

        # Refine: find offset of the actual anchor_text within the matched paragraph
        target_para = target_paragraphs[match.target_paragraph_index]
        anchor_offset = find_anchor_offset(comment.anchor_text, target_para)

        # Preserve cross-paragraph span relative to the matched start paragraph
        span = comment.end_paragraph_index - comment.start_paragraph_index
        target_end = min(
            match.target_paragraph_index + span,
            len(target_paragraphs) - 1,
        ) if span > 0 else None

        results.append(MatchResult(
            comment=comment,
            target_paragraph_index=match.target_paragraph_index,
            target_end_paragraph_index=target_end,
            score=match.score,
            anchor_offset=anchor_offset,
            below_threshold=match.score < threshold,
        ))

    return results
