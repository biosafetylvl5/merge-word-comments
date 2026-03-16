"""Orchestrate the merge pipeline: extract, match, insert."""

import copy
import io
import json
import time
from dataclasses import asdict
from pathlib import Path
from typing import Optional

from docx import Document
from lxml import etree
from rich.console import Console
from rich.table import Table

from merge_word_comments.extract import (
    extract_comments,
    extract_tracked_changes,
    build_heading_sections,
    _get_paragraph_texts,
)
from merge_word_comments.insert import (
    ensure_comments_part,
    insert_comments,
)
from merge_word_comments.match import (
    match_comments_to_target,
    find_best_paragraph_match,
    find_anchor_offset,
)
from merge_word_comments.types import Comment, FailureRecord, MatchResult, TrackedChange


WP_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{WP_NS}}}"

console = Console()


def _format_duration(seconds: float) -> str:
    """Format a duration in seconds for display."""
    if seconds < 60:
        return f"{seconds:.1f}s"
    minutes = int(seconds // 60)
    secs = seconds % 60
    return f"{minutes}m {secs:.0f}s"


def _get_target_paragraph_texts(doc: Document) -> list[str]:
    """Get plain text of each paragraph in the target document.

    Uses the same XML-based extraction as extract.py so that fuzzy
    matching compares like-for-like text representations.
    """
    return _get_paragraph_texts(doc.element.body)


def _recompute_anchor_offsets(
    match_results: list[MatchResult],
    final_paragraphs: list[str],
) -> list[MatchResult]:
    """Recompute anchor offsets against current paragraph texts."""
    refreshed: list[MatchResult] = []
    for m in match_results:
        if (
            not m.below_threshold
            and m.comment.anchor_text
            and m.target_paragraph_index < len(final_paragraphs)
        ):
            new_offset = find_anchor_offset(
                m.comment.anchor_text,
                final_paragraphs[m.target_paragraph_index],
            )
            refreshed.append(MatchResult(
                comment=m.comment,
                target_paragraph_index=m.target_paragraph_index,
                target_end_paragraph_index=m.target_end_paragraph_index,
                score=m.score,
                anchor_offset=new_offset,
                below_threshold=m.below_threshold,
            ))
        else:
            refreshed.append(m)
    return refreshed


def _intermediate_path(output_path: Path, step: int, total: int) -> Path:
    """Build an intermediate output path like merged.step-1of3.docx."""
    return output_path.parent / f"{output_path.stem}.step-{step}of{total}{output_path.suffix}"


def _save_intermediate(
    doc: Document,
    match_results: list[MatchResult],
    intermediate_path: Path,
) -> None:
    """Save an intermediate document with comments applied so far."""
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    intermediate_doc = Document(buf)

    final_paragraphs = _get_target_paragraph_texts(intermediate_doc)
    refreshed = _recompute_anchor_offsets(match_results, final_paragraphs)
    insert_comments(intermediate_doc, refreshed, intermediate_path)


def _print_failure_report(failures: list[FailureRecord]) -> None:
    """Print a detailed failure report to the console."""
    if not failures:
        return

    console.print()
    console.print(f"[bold red]Failed to merge {len(failures)} item(s)[/bold red]")
    console.print("These must be manually addressed:")
    console.print()

    # Group by source file
    by_file: dict[str, list[FailureRecord]] = {}
    for f in failures:
        by_file.setdefault(f.source_file, []).append(f)

    for source_file, file_failures in by_file.items():
        table = Table(
            title=f"Failures from {source_file}",
            show_lines=True,
            expand=True,
        )
        table.add_column("#", style="dim", width=3)
        table.add_column("Type", width=14)
        table.add_column("Author", width=16)
        table.add_column("Content", min_width=20)
        table.add_column("Source Location", width=18)
        table.add_column("Reason", min_width=24)

        for i, f in enumerate(file_failures, 1):
            if f.kind == "comment":
                type_str = "Comment"
            else:
                type_str = f"Tracked {f.change_type or 'change'}"

            location = f"para {f.source_paragraph_index}"
            if f.char_offset is not None:
                location += f", offset {f.char_offset}"

            reason_parts = [f.reason]
            if f.best_match_score is not None:
                reason_parts.append(
                    f"Best match: para {f.best_match_paragraph_index} "
                    f"(score {f.best_match_score:.1f})"
                )
            if f.best_match_paragraph_preview:
                reason_parts.append(
                    f"Target text: '{f.best_match_paragraph_preview}'"
                )
            if f.anchor_text:
                reason_parts.append(f"Anchor: '{f.anchor_text}'")

            table.add_row(
                str(i),
                type_str,
                f"{f.author}" + (f"\n{f.date}" if f.date else ""),
                f.content_preview,
                location,
                "\n".join(reason_parts),
            )

        console.print(table)
        console.print()


def _write_failure_json(failures: list[FailureRecord], output_path: Path) -> Path:
    """Write failures to a JSON file alongside the output."""
    json_path = output_path.parent / f"{output_path.stem}.failures.json"
    records = [asdict(f) for f in failures]
    json_path.write_text(json.dumps(records, indent=2, default=str))
    return json_path


def _relocate_comment_markers(
    tracked_change_el: etree._Element,
    paragraph_el: etree._Element,
) -> None:
    """Move comment range markers from inside a tracked change to adjacent positions.

    Tracked changes (w:ins/w:del) can contain nested commentRangeStart,
    commentRangeEnd, and commentReference elements from the source document.
    If left nested, they create duplicate markers when insert_comments() adds
    its own. This function relocates them to sibling positions in the paragraph:
    - commentRangeStart -> moved to just before the tracked change
    - commentRangeEnd and commentReference runs -> moved to just after
    """
    tc_index = list(paragraph_el).index(tracked_change_el)

    before_markers: list[etree._Element] = []
    after_markers: list[etree._Element] = []

    for el in list(tracked_change_el.iter()):
        if el is tracked_change_el:
            continue
        if el.tag == f"{W}commentRangeStart":
            el.getparent().remove(el)
            before_markers.append(el)
        elif el.tag == f"{W}commentRangeEnd":
            el.getparent().remove(el)
            after_markers.append(el)
        elif el.tag == f"{W}r":
            # Runs that exist solely to hold a commentReference
            refs = [c for c in el if c.tag == f"{W}commentReference"]
            if refs and not el.findall(f"{W}t"):
                el.getparent().remove(el)
                after_markers.append(el)

    # Insert commentRangeStart elements just before the tracked change
    for i, marker in enumerate(before_markers):
        paragraph_el.insert(tc_index + i, marker)

    # Insert commentRangeEnd / commentReference just after the tracked change
    after_start = tc_index + len(before_markers) + 1
    for i, marker in enumerate(after_markers):
        paragraph_el.insert(after_start + i, marker)


def _insert_change_at_offset(
    paragraph_el: etree._Element,
    change_el: etree._Element,
    char_offset: Optional[int],
) -> None:
    """Insert a tracked-change element at the correct character offset in a paragraph.

    If *char_offset* is None the element is appended at the end (fallback).
    Otherwise we walk w:r children, accumulate text lengths, and insert
    the change element at the position that corresponds to *char_offset*.
    """
    if char_offset is None:
        paragraph_el.append(change_el)
        return

    current_pos = 0
    children = list(paragraph_el)
    for i, child in enumerate(children):
        if child.tag != f"{W}r":
            continue
        run_len = 0
        for t in child.findall(f"{W}t"):
            run_len += len(t.text or "")
        if current_pos + run_len > char_offset:
            # The offset falls inside this run — split if needed
            offset_in_run = char_offset - current_pos
            if offset_in_run > 0:
                from merge_word_comments.insert import split_run_at_offset

                split_run_at_offset(child, offset_in_run)
                # After splitting, the new "after" run is the next sibling.
                # Insert the change element between the two halves.
                run_idx = list(paragraph_el).index(child)
                paragraph_el.insert(run_idx + 1, change_el)
            else:
                paragraph_el.insert(list(paragraph_el).index(child), change_el)
            return
        current_pos += run_len

    # Offset is past all runs — append at the end
    paragraph_el.append(change_el)


def _get_visible_text(paragraph_el: etree._Element) -> str:
    """Collect visible (w:t) text from direct w:r children of a paragraph."""
    parts: list[str] = []
    for child in paragraph_el:
        if child.tag != f"{W}r":
            continue
        for t in child.findall(f"{W}t"):
            parts.append(t.text or "")
    return "".join(parts)


def _find_closest_occurrence(
    text: str, target: str, hint_offset: int,
) -> int:
    """Find the occurrence of *target* in *text* closest to *hint_offset*.

    Returns the found offset, or -1 if *target* is not in *text*.
    """
    best_pos = -1
    best_dist = float("inf")
    start = 0
    while True:
        pos = text.find(target, start)
        if pos == -1:
            break
        dist = abs(pos - hint_offset)
        if dist < best_dist:
            best_dist = dist
            best_pos = pos
        start = pos + 1
    return best_pos


def _apply_deletion_at_offset(
    paragraph_el: etree._Element,
    change_el: etree._Element,
    char_offset: Optional[int],
    deletion_length: int,
    deletion_content: str = "",
) -> None:
    """Replace existing text with a w:del element at the correct offset.

    Uses content-based search: instead of trusting the pre-computed
    *char_offset*, searches for *deletion_content* in the paragraph's
    current visible text and uses the found position.  This handles
    offset drift (prior changes shifted text), duplicate deletions
    (text already in w:delText), and content mismatches (updated
    paragraph has different text).

    If *char_offset* is None the element is appended at the end (fallback).
    """
    from merge_word_comments.insert import split_run_at_offset

    if char_offset is None or deletion_length <= 0:
        paragraph_el.append(change_el)
        return

    # Content-based search: find where the deletion text actually is
    if deletion_content:
        visible = _get_visible_text(paragraph_el)
        found_offset = _find_closest_occurrence(
            visible, deletion_content, char_offset,
        )
        if found_offset >= 0:
            char_offset = found_offset
        else:
            # Content not in visible text — already deleted or text changed.
            # Insert the w:del element non-destructively (no text removal).
            _insert_change_at_offset(paragraph_el, change_el, char_offset)
            return

    # Phase 1: split at the start boundary so the deletion begins on a
    # run boundary.
    current_pos = 0
    for child in list(paragraph_el):
        if child.tag != f"{W}r":
            continue
        run_len = 0
        for t in child.findall(f"{W}t"):
            run_len += len(t.text or "")
        if current_pos + run_len > char_offset:
            offset_in_run = char_offset - current_pos
            if offset_in_run > 0:
                split_run_at_offset(child, offset_in_run)
            break
        current_pos += run_len

    # Phase 2: find the insertion index, then remove runs that the
    # deletion covers.  The change_el already contains the deleted text
    # as w:delText (cloned from the source document), so we just need to
    # remove the corresponding visible text from the paragraph.
    current_pos = 0
    insert_idx = None
    chars_remaining = deletion_length
    for child in list(paragraph_el):
        if child.tag != f"{W}r":
            continue
        run_len = 0
        for t in child.findall(f"{W}t"):
            run_len += len(t.text or "")
        if current_pos < char_offset:
            current_pos += run_len
            continue

        # This run is at or past the deletion start
        if insert_idx is None:
            insert_idx = list(paragraph_el).index(child)

        if run_len <= chars_remaining:
            # Entire run is consumed by the deletion — remove it
            chars_remaining -= run_len
            paragraph_el.remove(child)
            if chars_remaining == 0:
                break
        else:
            # Partial overlap — split and remove the first part
            split_run_at_offset(child, chars_remaining)
            paragraph_el.remove(child)
            chars_remaining = 0
            break
        current_pos += run_len

    # Phase 3: insert the w:del element where the removed text was
    if insert_idx is not None:
        paragraph_el.insert(insert_idx, change_el)
    else:
        paragraph_el.append(change_el)


def _apply_tracked_changes(
    doc: Document,
    changes: list[TrackedChange],
    target_paragraphs: list[str],
    threshold: int,
    verbose: bool = False,
    failures: Optional[list[FailureRecord]] = None,
    source_file: str = "",
    adaptive: bool = True,
    target_heading_sections: Optional[list[tuple[str, int, int]]] = None,
) -> None:
    """Insert tracked changes into the target document, respecting the threshold.

    Changes whose paragraph context scores below threshold are skipped — they
    almost certainly refer to content that was deleted from the updated doc.
    """
    body = doc.element.body
    para_elements = list(body.iter(f"{W}p"))

    tc_matched = 0
    tc_skipped = 0
    tc_no_match = 0
    total_changes = len(changes)
    change_times: list[float] = []

    for change_idx, change in enumerate(changes):
        change_start = time.monotonic()

        # Try the current (post-change) text first since it's closer to what
        # the updated document contains.  Fall back to the original (pre-change)
        # text if the current text scores worse.
        primary_context = change.paragraph_context_current or change.paragraph_context
        match = find_best_paragraph_match(
            primary_context, target_paragraphs, threshold,
            source_paragraph_index=change.paragraph_index,
            adaptive=adaptive,
            source_heading=change.source_heading,
            target_heading_sections=target_heading_sections,
        )

        # Fallback 1: if the current text scored below threshold, retry with
        # the original (pre-change) text which may preserve more context.
        if (
            match is not None
            and match.below_threshold
            and change.paragraph_context
            and change.paragraph_context != primary_context
        ):
            fallback = find_best_paragraph_match(
                change.paragraph_context, target_paragraphs, threshold,
                source_paragraph_index=change.paragraph_index,
                adaptive=adaptive,
                source_heading=change.source_heading,
                target_heading_sections=target_heading_sections,
            )
            if fallback is not None and fallback.score > match.score:
                match = fallback

        # Fallback 2: try sub-paragraph local context (~100 chars around
        # the change offset) which can match even when the full paragraph
        # context has diverged too much.
        if (
            match is not None
            and match.below_threshold
            and change.local_context
        ):
            local_match = find_best_paragraph_match(
                change.local_context, target_paragraphs, threshold,
                source_paragraph_index=change.paragraph_index,
                adaptive=adaptive,
                source_heading=change.source_heading,
                target_heading_sections=target_heading_sections,
            )
            if local_match is not None and local_match.score > match.score:
                match = local_match

        if match is None:
            tc_no_match += 1
            if failures is not None:
                failures.append(FailureRecord(
                    kind="tracked_change",
                    source_file=source_file,
                    author=change.author,
                    date=change.date,
                    content_preview=change.content,
                    reason="No matching paragraph found in target document",
                    source_paragraph_index=change.paragraph_index,
                    char_offset=change.char_offset,
                    anchor_text=change.paragraph_context,
                    change_type=change.change_type,
                    threshold=threshold,
                ))
            if verbose:
                console.print(
                    f"  [red]Tracked change no match[/red] "
                    f"'{change.content}' "
                    f"by {change.author} (no target paragraphs available)"
                )
        elif match.below_threshold:
            label = "insertion" if change.change_type == "insert" else "deletion"
            tc_skipped += 1
            target_preview = ""
            if match.target_paragraph_index < len(target_paragraphs):
                target_preview = target_paragraphs[match.target_paragraph_index]
            if failures is not None:
                failures.append(FailureRecord(
                    kind="tracked_change",
                    source_file=source_file,
                    author=change.author,
                    date=change.date,
                    content_preview=change.content,
                    reason=f"Match score {match.score:.1f} below threshold {threshold}",
                    source_paragraph_index=change.paragraph_index,
                    char_offset=change.char_offset,
                    anchor_text=change.paragraph_context,
                    change_type=change.change_type,
                    best_match_score=match.score,
                    best_match_paragraph_index=match.target_paragraph_index,
                    best_match_paragraph_preview=target_preview,
                    threshold=threshold,
                ))
            if verbose:
                console.print(
                    f"  [yellow]Tracked {label} skipped[/yellow] "
                    f"'{change.content}' "
                    f"by {change.author}"
                )
                console.print(
                    f"    score: {match.score:.1f} below threshold {threshold}, "
                    f"source para: {change.paragraph_index}, "
                    f"offset: {change.char_offset}"
                )
                console.print(
                    f"    context: '{change.paragraph_context}'"
                )
                tgt_preview = target_paragraphs[match.target_paragraph_index] if match.target_paragraph_index < len(target_paragraphs) else "<out of range>"
                console.print(
                    f"    best target para {match.target_paragraph_index}: "
                    f"'{tgt_preview}'"
                )
        else:
            label = "insertion" if change.change_type == "insert" else "deletion"
            tc_matched += 1
            target_idx = match.target_paragraph_index

            # Adjust target index when context came from a neighboring paragraph
            if change.context_paragraph_offset != 0:
                adjusted = target_idx - change.context_paragraph_offset
                adjusted = max(0, min(adjusted, len(para_elements) - 1))
                if verbose:
                    console.print(
                        f"  Tracked {label} -> paragraph {adjusted} "
                        f"(score: {match.score:.1f}, source para: {change.paragraph_index}, "
                        f"offset: {change.char_offset}, by {change.author}, "
                        f"matched neighbor para {target_idx}, adjusted by {-change.context_paragraph_offset:+d}) "
                        f"'{change.content}'"
                    )
                target_idx = adjusted
            elif verbose:
                console.print(
                    f"  Tracked {label} -> paragraph {target_idx} "
                    f"(score: {match.score:.1f}, source para: {change.paragraph_index}, "
                    f"offset: {change.char_offset}, by {change.author}) "
                    f"'{change.content}'"
                )

            if target_idx < len(para_elements):
                target_para = para_elements[target_idx]
                for xml_el in change.xml_elements:
                    cloned = copy.deepcopy(xml_el)
                    if change.change_type == "delete":
                        _apply_deletion_at_offset(
                            target_para, cloned, change.char_offset,
                            len(change.content), change.content,
                        )
                    else:
                        _insert_change_at_offset(target_para, cloned, change.char_offset)
                    # Relocate any comment markers that were nested inside the
                    # tracked change to adjacent positions in the paragraph.
                    _relocate_comment_markers(cloned, target_para)

        # Per-change time estimate (always runs regardless of outcome)
        change_elapsed = time.monotonic() - change_start
        change_times.append(change_elapsed)
        remaining_changes = total_changes - (change_idx + 1)
        if remaining_changes > 0:
            avg_per_change = sum(change_times) / len(change_times)
            est_remaining = avg_per_change * remaining_changes
            console.print(
                f"    [{change_idx + 1}/{total_changes}] "
                f"~{_format_duration(est_remaining)} remaining "
                f"for {remaining_changes} tracked change(s)"
            )

    if verbose:
        console.print(
            f"  Tracked change results: {tc_matched} matched, "
            f"{tc_skipped} skipped, {tc_no_match} unmatched"
        )


def merge_comments(
    updated_path: Path,
    original_paths: list[Path],
    output_path: Path,
    threshold: int = 80,
    verbose: bool = False,
    intermediates: bool = False,
    adaptive: bool = True,
) -> None:
    """Merge comments and tracked changes from originals into the updated document.

    Pipeline:
    1. Load the updated document as the base
    2. For each original:
       a. Extract + match comments
       b. Extract + apply tracked changes (inline, so verbose output is per-file)
       c. Optionally save intermediate document
    3. Insert all matched comments
    4. Save the result
    5. Print failure report and write failures JSON if any
    """
    start_time = time.monotonic()

    doc = Document(str(updated_path))
    target_paragraphs = _get_target_paragraph_texts(doc)
    target_heading_sections = build_heading_sections(doc.element.body)

    if verbose:
        console.print(f"Target document: {updated_path.name}")
        console.print(f"  {len(target_paragraphs)} paragraph(s) in target")
        non_empty = sum(1 for p in target_paragraphs if p.strip())
        console.print(f"  {non_empty} non-empty, {len(target_paragraphs) - non_empty} empty/whitespace-only")
        console.print(f"  Merging from {len(original_paths)} source file(s)")
        console.print(f"  Threshold: {threshold}")
        console.print()

    all_match_results: list[MatchResult] = []
    failures: list[FailureRecord] = []
    total_originals = len(original_paths)

    # Timing accumulators for per-item estimates
    total_comment_time = 0.0
    total_comment_count = 0
    total_tc_time = 0.0
    total_tc_count = 0

    for file_idx, orig_path in enumerate(original_paths):
        orig_start = time.monotonic()

        if verbose:
            console.print(f"Processing: {orig_path.name}")

        # --- Pre-processing time estimate (after first file) ---
        if file_idx > 0 and (total_comment_count > 0 or total_tc_count > 0):
            # Extract counts for this file to estimate time
            pre_comments = extract_comments(orig_path)
            pre_changes = extract_tracked_changes(orig_path)

            est_seconds = 0.0
            if total_comment_count > 0:
                est_seconds += (total_comment_time / total_comment_count) * len(pre_comments)
            if total_tc_count > 0:
                est_seconds += (total_tc_time / total_tc_count) * len(pre_changes)

            console.print(
                f"  {len(pre_comments)} comment(s), "
                f"{len(pre_changes)} tracked change(s) "
                f"— est. ~{_format_duration(est_seconds)}"
            )

            # Reuse the already-extracted data
            comments = pre_comments
            changes = pre_changes
        else:
            comments = None
            changes = None

        # --- Extract and match comments ---
        comment_start = time.monotonic()
        if comments is None:
            comments = extract_comments(orig_path)
        if verbose:
            authors = set(c.author for c in comments)
            console.print(
                f"  Found {len(comments)} comment(s)"
                + (f" by {', '.join(sorted(authors))}" if authors else "")
            )

        matches = match_comments_to_target(
            comments, target_paragraphs, threshold,
            adaptive=adaptive, target_heading_sections=target_heading_sections,
        )
        comment_elapsed = time.monotonic() - comment_start
        total_comment_time += comment_elapsed
        total_comment_count += max(len(comments), 1)  # avoid div-by-zero

        # Per-comment matching time (batch was done above; compute average)
        per_comment_avg = comment_elapsed / max(len(matches), 1)

        matched_count = 0
        skipped_count = 0
        total_matches = len(matches)
        for match_idx, m in enumerate(matches):
            if m.below_threshold:
                skipped_count += 1
                # Collect failure
                target_preview = ""
                if m.target_paragraph_index < len(target_paragraphs):
                    target_preview = target_paragraphs[m.target_paragraph_index]
                failures.append(FailureRecord(
                    kind="comment",
                    source_file=orig_path.name,
                    author=m.comment.author,
                    date=m.comment.date,
                    content_preview=m.comment.text,
                    reason=f"Match score {m.score:.1f} below threshold {threshold}",
                    source_paragraph_index=m.comment.start_paragraph_index,
                    anchor_text=m.comment.anchor_text,
                    anchor_context=m.comment.anchor_context,
                    best_match_score=m.score,
                    best_match_paragraph_index=m.target_paragraph_index,
                    best_match_paragraph_preview=target_preview,
                    threshold=threshold,
                ))
                if verbose:
                    console.print(
                        f"  [yellow]Comment skipped[/yellow] "
                        f"'{m.comment.text}' "
                        f"by {m.comment.author}"
                    )
                    console.print(
                        f"    score: {m.score:.1f} below threshold {threshold}, "
                        f"source para: {m.comment.start_paragraph_index}"
                    )
                    console.print(
                        f"    anchor: '{m.comment.anchor_text}'"
                    )
                    console.print(
                        f"    context: '{m.comment.anchor_context}'"
                    )
                    tgt_preview = target_paragraphs[m.target_paragraph_index] if m.target_paragraph_index < len(target_paragraphs) else "<out of range>"
                    console.print(
                        f"    best target para {m.target_paragraph_index}: "
                        f"'{tgt_preview}'"
                    )
            else:
                matched_count += 1
                if verbose:
                    span = ""
                    if m.target_end_paragraph_index is not None:
                        span = f"-{m.target_end_paragraph_index}"
                    console.print(
                        f"  Comment '{m.comment.text}' -> "
                        f"paragraph {m.target_paragraph_index}{span} "
                        f"(score: {m.score:.1f}, anchor: '{m.comment.anchor_text}', "
                        f"by {m.comment.author})"
                    )
            all_match_results.append(m)

            # Per-comment time estimate
            remaining_comments = total_matches - (match_idx + 1)
            if remaining_comments > 0:
                est_remaining = per_comment_avg * remaining_comments
                console.print(
                    f"    [{match_idx + 1}/{total_matches}] "
                    f"~{_format_duration(est_remaining)} remaining "
                    f"for {remaining_comments} comment(s)"
                )

        if verbose:
            console.print(
                f"  Comment results: {matched_count} matched, "
                f"{skipped_count} skipped"
            )

        # --- Extract and apply tracked changes ---
        tc_start = time.monotonic()
        if changes is None:
            changes = extract_tracked_changes(orig_path)
        if verbose:
            tc_authors = set(c.author for c in changes)
            insertions = sum(1 for c in changes if c.change_type == "insert")
            deletions = sum(1 for c in changes if c.change_type == "delete")
            console.print(
                f"  Found {len(changes)} tracked change(s) "
                f"({insertions} insertions, {deletions} deletions)"
                + (f" by {', '.join(sorted(tc_authors))}" if tc_authors else "")
            )
        if changes:
            _apply_tracked_changes(
                doc, changes, target_paragraphs, threshold, verbose,
                failures=failures, source_file=orig_path.name,
                adaptive=adaptive,
                target_heading_sections=target_heading_sections,
            )
            # Note: we intentionally do NOT refresh target_paragraphs here.
            # Using a frozen snapshot of the clean target document for ALL
            # matching prevents cascading text corruption from applied tracked
            # changes (deletions can leave garbled text that degrades fuzzy
            # matching for subsequent source files).

        tc_elapsed = time.monotonic() - tc_start
        total_tc_time += tc_elapsed
        total_tc_count += max(len(changes), 1)  # avoid div-by-zero

        if verbose:
            console.print()

        # --- Timing summary for this original ---
        orig_elapsed = time.monotonic() - orig_start
        per_comment = comment_elapsed / max(len(comments), 1)
        per_tc = tc_elapsed / max(len(changes), 1)

        timing_parts = []
        if comments:
            timing_parts.append(f"{len(comments)} comments @ ~{per_comment:.3f}s each")
        if changes:
            timing_parts.append(f"{len(changes)} tracked changes @ ~{per_tc:.3f}s each")
        timing_detail = f" ({', '.join(timing_parts)})" if timing_parts else ""

        console.print(
            f"  Processed {orig_path.name} in "
            f"{_format_duration(orig_elapsed)}{timing_detail}"
        )

        remaining = total_originals - (file_idx + 1)
        if remaining > 0:
            avg_per_comment = total_comment_time / total_comment_count
            avg_per_tc = total_tc_time / total_tc_count
            # Rough estimate: assume remaining files have similar counts
            avg_comments_per_file = total_comment_count / (file_idx + 1)
            avg_tcs_per_file = total_tc_count / (file_idx + 1)
            est_remaining = (avg_per_comment * avg_comments_per_file + avg_per_tc * avg_tcs_per_file) * remaining
            console.print(
                f"  Estimated time remaining: ~{_format_duration(est_remaining)} "
                f"({remaining} original(s) left)"
            )

        # --- Save intermediate document ---
        if intermediates and total_originals >= 2 and file_idx < total_originals - 1:
            inter_path = _intermediate_path(output_path, file_idx + 1, total_originals)
            _save_intermediate(doc, all_match_results, inter_path)
            console.print(
                f"  Intermediate saved: {inter_path.name} "
                f"(after {orig_path.name})"
            )

        console.print()

    # Recompute anchor offsets against the current (post-tracked-change)
    # document state so comments land on the correct text even when tracked
    # changes shifted character positions.
    final_paragraphs = _get_target_paragraph_texts(doc)
    refreshed_results = _recompute_anchor_offsets(all_match_results, final_paragraphs)

    # Insert all comments after tracked changes are in place
    out_of_range = insert_comments(doc, refreshed_results, output_path)

    # Convert out-of-range skips to failure records
    for m in out_of_range:
        num_paras = len(final_paragraphs)
        failures.append(FailureRecord(
            kind="comment",
            source_file="(unknown)",
            author=m.comment.author,
            date=m.comment.date,
            content_preview=m.comment.text,
            reason=f"Target paragraph {m.target_paragraph_index} out of range "
                   f"(document has {num_paras} paragraphs)",
            source_paragraph_index=m.comment.start_paragraph_index,
            anchor_text=m.comment.anchor_text,
            anchor_context=m.comment.anchor_context,
            best_match_score=m.score,
            best_match_paragraph_index=m.target_paragraph_index,
            threshold=threshold,
        ))

    total_elapsed = time.monotonic() - start_time

    if verbose:
        matched = sum(1 for m in all_match_results if not m.below_threshold)
        skipped = sum(1 for m in all_match_results if m.below_threshold)
        console.print(f"[bold]Merge complete[/bold]")
        console.print(f"  Comments: {matched} inserted, {skipped} skipped")
        console.print(f"  Output: {output_path}")

    console.print(f"Total elapsed: {_format_duration(total_elapsed)}")

    # Failure report (always shown, not gated behind --verbose)
    if failures:
        _print_failure_report(failures)
        json_path = _write_failure_json(failures, output_path)
        console.print(f"Failure details written to: {json_path}")
