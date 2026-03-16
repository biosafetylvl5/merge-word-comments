"""Insert comments and tracked changes into a target Word document."""

import copy
from pathlib import Path
from typing import Optional

from docx import Document
from lxml import etree

from merge_word_comments.types import MatchResult


WP_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{WP_NS}}}"

NSMAP = {"w": WP_NS}


def ensure_comments_part(doc: Document) -> None:
    """Ensure the document has a comments XML part (python-docx always creates one)."""
    # python-docx always creates _comments_part, so this is a no-op
    _ = doc.part._comments_part


def get_next_comment_id(doc: Document) -> int:
    """Get the next available comment ID in the document."""
    comments_el = doc.part._comments_part.element
    existing_ids = set()
    for cel in comments_el.findall(f"{W}comment"):
        try:
            existing_ids.add(int(cel.get(f"{W}id")))
        except (TypeError, ValueError):
            continue

    return max(existing_ids) + 1 if existing_ids else 0


def _find_run_and_offset_for_char_position(
    paragraph: etree._Element, char_offset: int
) -> tuple[Optional[etree._Element], int]:
    """Find which run contains the given character offset, and the offset within it.

    Iterates all descendant w:r elements (including those inside w:ins / w:del)
    so that character counting stays consistent with _get_paragraph_texts.
    Only direct-child w:r elements are returned as split/insert targets; runs
    inside tracked-change wrappers are counted for positioning but skipped as
    candidates.
    """
    current_pos = 0
    runs = list(paragraph.iter(f"{W}r"))
    last_direct_run = None

    for run in runs:
        # Track the most recent direct-child run (valid insertion target)
        if run.getparent() is paragraph:
            last_direct_run = run

        for t_el in run.findall(f"{W}t"):
            text = t_el.text or ""
            if current_pos + len(text) > char_offset:
                # If the matching run is inside a w:ins/w:del wrapper, we
                # can't split it — return the next direct-child run instead.
                if run.getparent() is not paragraph:
                    # Find the first direct-child run AFTER this wrapper
                    wrapper = run.getparent()
                    siblings = list(paragraph)
                    wrapper_idx = siblings.index(wrapper)
                    for sibling in siblings[wrapper_idx + 1:]:
                        if sibling.tag == f"{W}r":
                            return sibling, 0
                    # No direct run after the wrapper — return last direct run
                    if last_direct_run is not None:
                        return last_direct_run, -1
                    return None, 0
                return run, char_offset - current_pos
            current_pos += len(text)

    # Past end: return last direct-child run
    if last_direct_run is not None:
        return last_direct_run, -1

    return None, 0


def split_run_at_offset(run, offset: int):
    """Split a python-docx Run at a character offset, returning (before, after) runs.

    Both runs preserve the original formatting.
    """
    from docx.text.run import Run

    # Get the run element (lxml)
    if isinstance(run, Run):
        run_el = run._element
        parent = run_el.getparent()
    else:
        run_el = run
        parent = run_el.getparent()

    # Get the text
    full_text = ""
    t_elements = run_el.findall(f"{W}t")
    for t in t_elements:
        full_text += t.text or ""

    before_text = full_text[:offset]
    after_text = full_text[offset:]

    # Create the "after" run as a copy
    after_run_el = copy.deepcopy(run_el)

    # Set text on original (becomes "before")
    for t in run_el.findall(f"{W}t"):
        run_el.remove(t)
    t_before = etree.SubElement(run_el, f"{W}t")
    t_before.text = before_text
    if before_text.startswith(" ") or before_text.endswith(" "):
        t_before.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    # Set text on copy (becomes "after")
    for t in after_run_el.findall(f"{W}t"):
        after_run_el.remove(t)
    t_after = etree.SubElement(after_run_el, f"{W}t")
    t_after.text = after_text
    if after_text.startswith(" ") or after_text.endswith(" "):
        t_after.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    # Insert after_run right after the original run
    parent.insert(list(parent).index(run_el) + 1, after_run_el)

    # Always return lxml elements. Constructing new python-docx Run
    # objects from raw elements is fragile (Run.__init__ expects proxy
    # parents), and all internal callers work with lxml elements.
    return run_el, after_run_el


def _create_comment_element(
    comment_id: int,
    author: str,
    initials: Optional[str],
    date: Optional[str],
    text: str,
) -> etree._Element:
    """Create a w:comment XML element."""
    attrs = {
        f"{W}id": str(comment_id),
        f"{W}author": author,
    }
    if initials:
        attrs[f"{W}initials"] = initials
    if date:
        attrs[f"{W}date"] = date

    comment_el = etree.Element(f"{W}comment", attrib=attrs, nsmap=NSMAP)

    # Add paragraph with comment text
    p_el = etree.SubElement(comment_el, f"{W}p")
    r_el = etree.SubElement(p_el, f"{W}r")
    t_el = etree.SubElement(r_el, f"{W}t")
    t_el.text = text

    return comment_el


def _make_ref_run(comment_id: int) -> etree._Element:
    """Create the w:r containing w:commentReference."""
    id_str = str(comment_id)
    ref_run = etree.Element(f"{W}r")
    ref_rpr = etree.SubElement(ref_run, f"{W}rPr")
    etree.SubElement(ref_rpr, f"{W}rStyle", attrib={f"{W}val": "CommentReference"})
    etree.SubElement(ref_run, f"{W}commentReference", attrib={f"{W}id": id_str})
    return ref_run


def _prepend_range_start(paragraph_el: etree._Element, id_str: str,
                          anchor_offset: Optional[int]) -> None:
    """Insert commentRangeStart at the anchor position (or beginning) of a paragraph."""
    range_start = etree.Element(f"{W}commentRangeStart", attrib={f"{W}id": id_str})

    if anchor_offset is not None and anchor_offset >= 0:
        run_el, offset_in_run = _find_run_and_offset_for_char_position(
            paragraph_el, anchor_offset
        )
        if run_el is not None:
            run_idx = list(paragraph_el).index(run_el)
            if offset_in_run > 0:
                split_run_at_offset(run_el, offset_in_run)
                paragraph_el.insert(run_idx + 1, range_start)
            else:
                paragraph_el.insert(run_idx, range_start)
            return

    # Fallback: insert before the first run
    for i, child in enumerate(paragraph_el):
        if child.tag == f"{W}r":
            paragraph_el.insert(i, range_start)
            return
    paragraph_el.append(range_start)


def _append_range_end(paragraph_el: etree._Element, id_str: str,
                       anchor_offset: Optional[int], anchor_length: int,
                       start_para: bool) -> None:
    """Insert commentRangeEnd + commentReference into a paragraph.

    When this is the end paragraph of a cross-para comment (start_para=False),
    offset is ignored and the markers go at the beginning of the paragraph.
    """
    range_end = etree.Element(f"{W}commentRangeEnd", attrib={f"{W}id": id_str})
    ref_run = _make_ref_run(int(id_str))

    if not start_para:
        # Cross-para: end markers go at the very start of this paragraph
        paragraph_el.insert(0, range_end)
        paragraph_el.insert(1, ref_run)
        return

    # Same paragraph: place end after the anchor text
    if anchor_offset is not None and anchor_offset >= 0:
        end_char = anchor_offset + anchor_length
        end_run, end_offset = _find_run_and_offset_for_char_position(
            paragraph_el, end_char
        )
        if end_run is not None:
            end_idx = list(paragraph_el).index(end_run)
            if end_offset > 0 and end_offset != -1:
                split_run_at_offset(end_run, end_offset)
                # After split, insert after the "before" half
                paragraph_el.insert(end_idx + 1, range_end)
                paragraph_el.insert(end_idx + 2, ref_run)
            elif anchor_length == 0:
                # Zero-length anchor: end marker goes at the same position
                # as the start marker (before this run).
                paragraph_el.insert(end_idx, range_end)
                paragraph_el.insert(end_idx + 1, ref_run)
            elif end_offset == 0:
                # End position is at the very start of this run (run boundary).
                # Insert BEFORE this run so the range doesn't include its text.
                paragraph_el.insert(end_idx, range_end)
                paragraph_el.insert(end_idx + 1, ref_run)
            else:
                # end_offset == -1 (past end of all text): insert after this run
                paragraph_el.insert(end_idx + 1, range_end)
                paragraph_el.insert(end_idx + 2, ref_run)
            return

    paragraph_el.append(range_end)
    paragraph_el.append(ref_run)


def _cleanup_orphaned_markers(doc: Document) -> list[str]:
    """Remove comment range markers whose IDs don't match any comment in comments.xml.

    Markers can end up orphaned when tracked changes carry embedded comment
    markers from a source document. Those markers are relocated to adjacent
    positions (preserving their location) but have stale IDs from the source.
    This cleanup removes them after all comments have been properly inserted.

    Returns a list of warning strings describing what was cleaned up.
    """
    warnings: list[str] = []
    comments_el = doc.part._comments_part.element
    body = doc.element.body

    comment_ids = {
        cel.get(f"{W}id")
        for cel in comments_el.findall(f"{W}comment")
    }

    # Also track duplicates — keep only the first occurrence of each ID
    from collections import Counter
    seen_start_ids: Counter[str] = Counter()
    seen_end_ids: Counter[str] = Counter()

    marker_tags = {f"{W}commentRangeStart", f"{W}commentRangeEnd"}

    for el in list(body.iter(f"{W}commentRangeStart")):
        cid = el.get(f"{W}id")
        seen_start_ids[cid] += 1
        if cid not in comment_ids:
            warnings.append(
                f"Removed orphaned commentRangeStart for non-existent comment {cid}"
            )
            el.getparent().remove(el)
        elif seen_start_ids[cid] > 1:
            warnings.append(
                f"Removed duplicate commentRangeStart for comment {cid}"
            )
            el.getparent().remove(el)

    for el in list(body.iter(f"{W}commentRangeEnd")):
        cid = el.get(f"{W}id")
        seen_end_ids[cid] += 1
        if cid not in comment_ids:
            warnings.append(
                f"Removed orphaned commentRangeEnd for non-existent comment {cid}"
            )
            el.getparent().remove(el)
        elif seen_end_ids[cid] > 1:
            warnings.append(
                f"Removed duplicate commentRangeEnd for comment {cid}"
            )
            el.getparent().remove(el)

    # Clean up orphaned and duplicate commentReference elements
    seen_ref_ids: Counter[str] = Counter()
    for el in list(body.iter(f"{W}commentReference")):
        cid = el.get(f"{W}id")
        seen_ref_ids[cid] += 1
        should_remove = False

        if cid not in comment_ids:
            warnings.append(
                f"Removed orphaned commentReference for non-existent comment {cid}"
            )
            should_remove = True
        elif seen_ref_ids[cid] > 1:
            warnings.append(
                f"Removed duplicate commentReference for comment {cid}"
            )
            should_remove = True

        if should_remove:
            run = el.getparent()
            if run is not None and run.tag == f"{W}r" and not run.findall(f"{W}t"):
                # Run exists solely to hold the commentReference — remove whole run
                run.getparent().remove(run)
            else:
                # Run has other content — remove just the commentReference element
                el.getparent().remove(el)

    return warnings


def insert_comments(
    doc: Document,
    match_results: list[MatchResult],
    output_path: Path,
) -> list[MatchResult]:
    """Insert comments into the document based on match results, then save.

    Returns a list of MatchResults that were skipped because their target
    paragraph index was out of range.
    """
    skipped: list[MatchResult] = []

    if not match_results:
        doc.save(str(output_path))
        return skipped

    ensure_comments_part(doc)
    next_id = get_next_comment_id(doc)

    comments_el = doc.part._comments_part.element
    body = doc.element.body
    paragraphs = list(body.iter(f"{W}p"))

    for match in match_results:
        if match.below_threshold:
            continue

        comment = match.comment

        start_idx = match.target_paragraph_index
        end_idx = match.effective_end_index()
        anchor_len = len(comment.anchor_text) if comment.anchor_text else 0

        if start_idx >= len(paragraphs):
            skipped.append(match)
            continue

        comment_id = next_id
        next_id += 1
        id_str = str(comment_id)

        # Add comment element to comments.xml, preserving original XML
        # (rich formatting, reply threads, etc.) when available.
        if comment.xml_element is not None:
            cel = copy.deepcopy(comment.xml_element)
            cel.set(f"{W}id", str(comment_id))
            comments_el.append(cel)
        else:
            comments_el.append(_create_comment_element(
                comment_id=comment_id,
                author=comment.author,
                initials=comment.initials,
                date=comment.date,
                text=comment.text,
            ))

        if start_idx == end_idx:
            # Single-paragraph comment: start and end go in the same paragraph
            _prepend_range_start(paragraphs[start_idx], id_str, match.anchor_offset)
            _append_range_end(paragraphs[start_idx], id_str, match.anchor_offset,
                              anchor_len, start_para=True)
        else:
            # Cross-paragraph: start in first para, end in last para
            _prepend_range_start(paragraphs[start_idx], id_str, match.anchor_offset)
            end_para_idx = min(end_idx, len(paragraphs) - 1)
            _append_range_end(paragraphs[end_para_idx], id_str, None, 0,
                              start_para=False)

    # Clean up any orphaned or duplicate markers (e.g. from tracked changes
    # that carried embedded comment markers from the source document).
    _cleanup_orphaned_markers(doc)

    doc.save(str(output_path))

    return skipped
