"""Extract comments and tracked changes from Word documents."""

import copy
import re
from pathlib import Path
from typing import Optional

from docx import Document
from lxml import etree

from merge_word_comments.types import Comment, TrackedChange


WP_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{WP_NS}}}"

MIN_CONTEXT_LENGTH = 80
MAX_CONTEXT_LENGTH = 200
_URL_PATTERN = re.compile(r"https?://\S+")


def _get_original_paragraph_texts(body: etree._Element) -> list[str]:
    """Extract paragraph text as it appeared before any tracked changes.

    Reconstructs the pre-change text for matching purposes:
    - w:r (direct child): include w:t text (normal text)
    - w:del (direct child): include w:delText (was in original)
    - w:ins (direct child): skip entirely (was NOT in original)
    """
    texts = []
    for para in body.iter(f"{W}p"):
        parts: list[str] = []
        for child in para:
            if child.tag == f"{W}ins":
                continue
            if child.tag == f"{W}del":
                for dt in child.iter(f"{W}delText"):
                    if dt.text:
                        parts.append(dt.text)
            elif child.tag == f"{W}r":
                for t in child.findall(f"{W}t"):
                    if t.text:
                        parts.append(t.text)
        texts.append("".join(parts))
    return texts


def _get_paragraph_texts(body: etree._Element) -> list[str]:
    """Extract plain text from each w:p element in the body.

    Finds all w:p descendants (including those inside tables) so that
    comments on table content can be matched and placed correctly.

    Only includes w:t text (not w:delText) so that offsets are consistent
    with _find_run_and_offset_for_char_position, which also counts only
    w:t when navigating to a character position for comment insertion.
    """
    texts = []
    for para in body.iter(f"{W}p"):
        runs_text = []
        for run in para.iter(f"{W}r"):
            for t in run.findall(f"{W}t"):
                if t.text:
                    runs_text.append(t.text)
        texts.append("".join(runs_text))
    return texts


def _get_paragraph_elements(body: etree._Element) -> list[etree._Element]:
    """Get all w:p elements in the body, including those inside tables."""
    return list(body.iter(f"{W}p"))


def _extract_heading_keyword(heading_text: str) -> str:
    """Extract the entry keyword from a heading paragraph.

    Headings look like "ANDROGYNY. The concept of..." — we extract just
    the keyword before the first period (e.g. "ANDROGYNY").
    """
    if not heading_text:
        return ""
    dot_idx = heading_text.find(".")
    if dot_idx > 0:
        return heading_text[:dot_idx].strip()
    # Fallback: first two words
    words = heading_text.split()[:2]
    return " ".join(words).strip()


def build_heading_sections(
    body: etree._Element,
) -> list[tuple[str, int, int]]:
    """Build a list of (heading_keyword, start_para_idx, end_para_idx) sections.

    Scans all paragraphs for Heading styles and extracts the entry keyword
    (first 1-2 words before the period).
    """
    paragraphs = list(body.iter(f"{W}p"))
    total = len(paragraphs)
    headings: list[tuple[str, int]] = []

    for i, para in enumerate(paragraphs):
        pPr = para.find(f"{W}pPr")
        if pPr is None:
            continue
        pStyle = pPr.find(f"{W}pStyle")
        if pStyle is None:
            continue
        style_val = pStyle.get(f"{W}val", "")
        if "Heading" not in style_val:
            continue
        text = "".join(t.text or "" for t in para.iter(f"{W}t"))
        keyword = _extract_heading_keyword(text)
        if keyword:
            headings.append((keyword, i))

    sections: list[tuple[str, int, int]] = []
    for j, (kw, start) in enumerate(headings):
        end = headings[j + 1][1] if j + 1 < len(headings) else total
        sections.append((kw, start, end))
    return sections


def find_enclosing_heading(
    para_index: int,
    sections: list[tuple[str, int, int]],
) -> str:
    """Find the heading keyword for the section containing para_index."""
    for keyword, start, end in sections:
        if start <= para_index < end:
            return keyword
    return ""


def _find_paragraph_index(element: etree._Element,
                          paragraphs: list[etree._Element]) -> int:
    """Find which paragraph index an element belongs to."""
    node = element
    while node is not None:
        if node.tag == f"{W}p" and node in paragraphs:
            return paragraphs.index(node)
        node = node.getparent()
    # If not found inside a paragraph, find nearest preceding paragraph
    body = paragraphs[0].getparent() if paragraphs else None
    if body is not None:
        for i, para in enumerate(paragraphs):
            # Check if element comes before this paragraph in document order
            if element.getparent() is body:
                body_children = list(body)
                try:
                    elem_idx = body_children.index(element)
                    para_idx = body_children.index(para)
                    if elem_idx <= para_idx:
                        return i
                except ValueError:
                    continue
    return 0


def _collect_text_in_range(
    start_el: etree._Element,
    end_el: etree._Element,
    body: etree._Element,
) -> str:
    """Collect all text between commentRangeStart and commentRangeEnd."""
    collecting = False
    parts = []

    for element in body.iter():
        if element is start_el:
            collecting = True
            continue
        if element is end_el:
            break
        if collecting and element.tag in (f"{W}t", f"{W}delText") and element.text:
            parts.append(element.text)

    return "".join(parts)


def _build_anchor_context(
    anchor_text: str,
    paragraph_texts: list[str],
    para_index: int,
) -> str:
    """Expand short anchor text with surrounding paragraph context."""
    # If anchor is primarily a URL, use paragraph context instead since
    # URLs are often removed or reformatted in the updated document.
    cleaned = _URL_PATTERN.sub("", anchor_text).strip("() ")
    if len(cleaned) < 10 and len(anchor_text) > len(cleaned):
        if 0 <= para_index < len(paragraph_texts):
            return paragraph_texts[para_index][:200]

    if len(anchor_text) >= MIN_CONTEXT_LENGTH:
        return anchor_text

    if 0 <= para_index < len(paragraph_texts):
        para_text = paragraph_texts[para_index]
        if not anchor_text or anchor_text not in para_text:
            return para_text

        # Return a focused character window around the anchor rather than
        # the entire paragraph, so the context doesn't dilute matching.
        idx = para_text.find(anchor_text)
        half_extra = (MIN_CONTEXT_LENGTH - len(anchor_text)) // 2
        start = max(0, idx - half_extra)
        end = min(len(para_text), idx + len(anchor_text) + half_extra)
        needed = MIN_CONTEXT_LENGTH - (end - start)
        if needed > 0:
            if start == 0:
                end = min(len(para_text), end + needed)
            else:
                start = max(0, start - needed)
        return para_text[start:end]

    return anchor_text


def extract_comments(doc_path: Path) -> list[Comment]:
    """Extract all comments from a Word document with their anchor text."""
    doc = Document(str(doc_path))
    body = doc.element.body
    paragraphs = _get_paragraph_elements(body)
    paragraph_texts = _get_paragraph_texts(body)
    heading_sections = build_heading_sections(body)

    # Get comments part (python-docx always creates one)
    comments_part = doc.part._comments_part
    comments_element = comments_part.element
    comment_elements = comments_element.findall(f"{W}comment")

    if not comment_elements:
        return []

    # Build a map of comment_id -> comment XML element
    comment_map = {}
    for cel in comment_elements:
        cid = int(cel.get(f"{W}id"))
        comment_map[cid] = cel

    # Find all commentRangeStart/End pairs in the body
    range_starts = {}
    range_ends = {}
    for el in body.iter():
        if el.tag == f"{W}commentRangeStart":
            cid = int(el.get(f"{W}id"))
            range_starts[cid] = el
        elif el.tag == f"{W}commentRangeEnd":
            cid = int(el.get(f"{W}id"))
            range_ends[cid] = el

    results = []
    for cid, cel in comment_map.items():
        # Extract comment text
        text_parts = []
        for t in cel.iter(f"{W}t"):
            if t.text:
                text_parts.append(t.text)
        comment_text = "".join(text_parts)

        author = cel.get(f"{W}author", "")
        initials = cel.get(f"{W}initials")
        date = cel.get(f"{W}date")

        # Get anchor text from range markers
        start_el = range_starts.get(cid)
        end_el = range_ends.get(cid)

        anchor_text = ""
        start_para_idx = 0
        end_para_idx = 0

        if start_el is not None and end_el is not None:
            anchor_text = _collect_text_in_range(start_el, end_el, body)
            start_para_idx = _find_paragraph_index(start_el, paragraphs)
            end_para_idx = _find_paragraph_index(end_el, paragraphs)
        elif start_el is not None:
            start_para_idx = _find_paragraph_index(start_el, paragraphs)
            end_para_idx = start_para_idx
        elif end_el is not None:
            end_para_idx = _find_paragraph_index(end_el, paragraphs)
            start_para_idx = end_para_idx

        anchor_context = _build_anchor_context(
            anchor_text, paragraph_texts, start_para_idx
        )

        results.append(Comment(
            comment_id=cid,
            author=author,
            initials=initials,
            date=date,
            text=comment_text,
            anchor_text=anchor_text,
            anchor_context=anchor_context,
            start_paragraph_index=start_para_idx,
            end_paragraph_index=end_para_idx,
            xml_element=copy.deepcopy(cel),
            source_heading=find_enclosing_heading(start_para_idx, heading_sections),
        ))

    return results


def _compute_change_char_offset(change_el: etree._Element) -> Optional[int]:
    """Compute the character offset of a tracked-change element within its parent paragraph.

    Sums the text length of all preceding sibling runs (w:r elements and
    other tracked-change elements) so the change can be repositioned at
    the correct location in a target paragraph.

    Also detects mid-word insertions: if the preceding w:r run ends with a
    word character and the following w:r run starts with one, the change
    element sits between two fragments of the same word (e.g. "discus" +
    w:ins + "sions" representing "discussions").  In that case the offset is
    advanced past the word suffix so the insertion lands after the complete
    word in the target paragraph instead of splitting it.
    """
    parent = change_el.getparent()
    if parent is None:
        return None

    offset = 0
    preceding_text = ""
    for sibling in parent:
        if sibling is change_el:
            break
        if sibling.tag == f"{W}r":
            # Regular run: count visible text.
            for t in sibling.findall(f"{W}t"):
                if t.text:
                    preceding_text += t.text
                    offset += len(t.text)
        elif sibling.tag == f"{W}del":
            # Deleted text IS present in the target paragraph (which represents
            # the original/pre-change document).  Counting it keeps char_offset
            # aligned with the target's character positions.
            for dt in sibling.iter(f"{W}delText"):
                if dt.text:
                    preceding_text += dt.text
                    offset += len(dt.text)
        # w:ins siblings: their text is NOT in the target (target = original,
        # before the source's tracked changes are applied), so skip them.

    # Mid-word detection: if the insertion falls between two word-character
    # fragments, advance the offset to the end of the word.  This avoids
    # splitting a word like "discussions" when "discus" and "sions" are in
    # separate source runs on either side of the tracked-change element.
    if preceding_text and preceding_text[-1].isalnum():
        found_change = False
        for sibling in parent:
            if found_change:
                if sibling.tag == f"{W}r":
                    following_text = "".join(
                        t.text or "" for t in sibling.findall(f"{W}t")
                    )
                    if following_text and following_text[0].isalnum():
                        for ch in following_text:
                            if ch.isalnum():
                                offset += 1
                            else:
                                break
                    break  # only look at the first following w:r sibling
            if sibling is change_el:
                found_change = True

    return offset


def _find_neighbor_context(
    para_idx: int,
    paragraph_texts: list[str],
) -> tuple[str, int]:
    """Find the nearest non-empty neighbor paragraph text when the current one is empty.

    Returns (context_text, offset) where offset is the distance to the neighbor:
    -1 = preceding paragraph, +1 = following, -2 = two before, etc.
    Returns ("", 0) if no non-empty neighbor is found.
    """
    max_distance = min(5, len(paragraph_texts))
    for distance in range(1, max_distance + 1):
        # Check preceding paragraph first
        prev_idx = para_idx - distance
        if 0 <= prev_idx < len(paragraph_texts) and paragraph_texts[prev_idx].strip():
            return paragraph_texts[prev_idx], -distance
        # Then following paragraph
        next_idx = para_idx + distance
        if 0 <= next_idx < len(paragraph_texts) and paragraph_texts[next_idx].strip():
            return paragraph_texts[next_idx], distance
    return "", 0


def extract_tracked_changes(doc_path: Path) -> list[TrackedChange]:
    """Extract all tracked changes (insertions/deletions) from a Word document."""
    doc = Document(str(doc_path))
    body = doc.element.body
    paragraphs = _get_paragraph_elements(body)
    original_texts = _get_original_paragraph_texts(body)
    current_texts = _get_paragraph_texts(body)
    heading_sections = build_heading_sections(body)

    results = []

    for change_tag, change_type in [(f"{W}ins", "insert"), (f"{W}del", "delete")]:
        for element in body.iter(change_tag):
            author = element.get(f"{W}author", "")
            date = element.get(f"{W}date")

            # Extract the content
            text_parts = []
            if change_type == "delete":
                for dt in element.iter(f"{W}delText"):
                    if dt.text:
                        text_parts.append(dt.text)
            else:
                for t in element.iter(f"{W}t"):
                    if t.text:
                        text_parts.append(t.text)

            content = "".join(text_parts)
            para_idx = _find_paragraph_index(element, paragraphs)
            para_context = original_texts[para_idx] if para_idx < len(original_texts) else ""
            para_context_current = current_texts[para_idx] if para_idx < len(current_texts) else ""

            # When the paragraph has no original text (entirely new content),
            # use the nearest non-empty neighbor paragraph as context so the
            # change can be located in the target document.
            context_offset = 0
            if not para_context.strip():
                neighbor_text, context_offset = _find_neighbor_context(
                    para_idx, original_texts,
                )
                if neighbor_text:
                    para_context = neighbor_text

            # Compute character offset within the paragraph by summing
            # text of all preceding sibling runs.
            char_offset = _compute_change_char_offset(element)

            # Build a tight local context (~100 chars) around the change
            # offset for sub-paragraph matching as a last-resort fallback.
            local_context = ""
            local_src = para_context_current or para_context
            if char_offset is not None and local_src:
                lc_half = 50
                lc_start = max(0, char_offset - lc_half)
                lc_end = min(len(local_src), char_offset + len(content) + lc_half)
                local_context = local_src[lc_start:lc_end]

            # Cap very long paragraph contexts to a window around the change
            # offset.  partial_ratio degrades when both query and candidate
            # are very long (800+ chars).
            if len(para_context) > MAX_CONTEXT_LENGTH and char_offset is not None:
                half = MAX_CONTEXT_LENGTH // 2
                ctx_start = max(0, char_offset - half)
                ctx_end = min(len(para_context), char_offset + half)
                # Give slack to the other side if one side hit a boundary
                needed = MAX_CONTEXT_LENGTH - (ctx_end - ctx_start)
                if needed > 0:
                    if ctx_start == 0:
                        ctx_end = min(len(para_context), ctx_end + needed)
                    else:
                        ctx_start = max(0, ctx_start - needed)
                para_context = para_context[ctx_start:ctx_end]
            if len(para_context_current) > MAX_CONTEXT_LENGTH and char_offset is not None:
                half = MAX_CONTEXT_LENGTH // 2
                ctx_start = max(0, char_offset - half)
                ctx_end = min(len(para_context_current), char_offset + half)
                needed = MAX_CONTEXT_LENGTH - (ctx_end - ctx_start)
                if needed > 0:
                    if ctx_start == 0:
                        ctx_end = min(len(para_context_current), ctx_end + needed)
                    else:
                        ctx_start = max(0, ctx_start - needed)
                para_context_current = para_context_current[ctx_start:ctx_end]

            results.append(TrackedChange(
                change_type=change_type,
                author=author,
                date=date,
                content=content,
                paragraph_context=para_context,
                paragraph_index=para_idx,
                char_offset=char_offset,
                context_paragraph_offset=context_offset,
                paragraph_context_current=para_context_current,
                local_context=local_context,
                source_heading=find_enclosing_heading(para_idx, heading_sections),
                xml_elements=[copy.deepcopy(element)],
            ))

    return results
