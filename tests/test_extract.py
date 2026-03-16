"""Tests for comment and tracked change extraction."""

from docx import Document

from merge_word_comments.extract import extract_comments, extract_tracked_changes
from merge_word_comments.types import Comment, TrackedChange


class TestExtractComments:
    """Tests for extracting comments from Word documents."""

    def test_no_comments_returns_empty_list(self, original_no_comments_path):
        comments = extract_comments(original_no_comments_path)
        assert comments == []

    def test_single_comment_extracted(self, original_with_comments_path):
        comments = extract_comments(original_with_comments_path)
        assert len(comments) == 1

    def test_single_comment_has_correct_text(self, original_with_comments_path):
        comments = extract_comments(original_with_comments_path)
        assert "comment on the word the" in comments[0].text

    def test_single_comment_has_author(self, original_with_comments_path):
        comments = extract_comments(original_with_comments_path)
        assert comments[0].author == "Uttmark,Gwyn"

    def test_single_comment_has_anchor_text(self, original_with_comments_path):
        comments = extract_comments(original_with_comments_path)
        assert "the" in comments[0].anchor_text.lower()

    def test_returns_comment_dataclass(self, original_with_comments_path):
        comments = extract_comments(original_with_comments_path)
        assert isinstance(comments[0], Comment)

    def test_multiple_comments_extracted(self, original_with_comments2_path):
        comments = extract_comments(original_with_comments2_path)
        assert len(comments) == 4

    def test_comment_on_single_word(self, original_with_comments2_path):
        """Comment on 'Snowball' should have that as anchor text."""
        comments = extract_comments(original_with_comments2_path)
        snowball_comments = [c for c in comments if "snowball" in c.text.lower()]
        assert len(snowball_comments) == 1
        assert "Snowball" in snowball_comments[0].anchor_text

    def test_cross_paragraph_comment(self, original_with_comments2_path):
        """Comment spanning two paragraphs should capture anchor text from both."""
        comments = extract_comments(original_with_comments2_path)
        cross_para = [c for c in comments if "across two paragraphs" in c.text]
        assert len(cross_para) == 1
        assert len(cross_para[0].anchor_text) > 0

    def test_cross_paragraph_comment_spans_multiple_paragraphs(
        self, original_with_comments2_path
    ):
        comments = extract_comments(original_with_comments2_path)
        cross_para = [c for c in comments if "across two paragraphs" in c.text]
        assert cross_para[0].start_paragraph_index != cross_para[0].end_paragraph_index

    def test_comment_has_date(self, original_with_comments_path):
        comments = extract_comments(original_with_comments_path)
        assert comments[0].date is not None

    def test_comment_with_context(self, original_with_comments_path):
        """Short anchor text should include surrounding context for matching."""
        comments = extract_comments(original_with_comments_path)
        # "the" is very short, context should include surrounding words
        assert len(comments[0].anchor_context) > len(comments[0].anchor_text)

    def test_comments_with_empty_anchor(self, original_with_comments3_path):
        """Comments on deleted/edited regions may have empty direct anchor text."""
        comments = extract_comments(original_with_comments3_path)
        assert len(comments) == 4

    def test_comment_paragraph_index_is_set(self, original_with_comments_path):
        comments = extract_comments(original_with_comments_path)
        assert comments[0].start_paragraph_index >= 0


class TestExtractTrackedChanges:
    """Tests for extracting tracked changes from Word documents."""

    def test_no_tracked_changes_returns_empty(self, original_no_comments_path):
        changes = extract_tracked_changes(original_no_comments_path)
        assert changes == []

    def test_no_tracked_changes_in_comments_only_doc(
        self, original_with_comments_path
    ):
        changes = extract_tracked_changes(original_with_comments_path)
        assert changes == []

    def test_tracked_changes_found_in_edited_doc(self, original_with_comments3_path):
        """original_with_comments3 has text edits (tracked changes)."""
        changes = extract_tracked_changes(original_with_comments3_path)
        assert len(changes) > 0

    def test_tracked_change_has_type(self, original_with_comments3_path):
        changes = extract_tracked_changes(original_with_comments3_path)
        assert all(c.change_type in ("insert", "delete") for c in changes)

    def test_tracked_change_has_author(self, original_with_comments3_path):
        changes = extract_tracked_changes(original_with_comments3_path)
        assert all(c.author is not None for c in changes)

    def test_tracked_change_has_content(self, original_with_comments3_path):
        """Each tracked change should carry the inserted or deleted text."""
        changes = extract_tracked_changes(original_with_comments3_path)
        assert all(c.content is not None for c in changes)

    def test_tracked_change_has_paragraph_context(self, original_with_comments3_path):
        changes = extract_tracked_changes(original_with_comments3_path)
        assert all(c.paragraph_context is not None for c in changes)

    def test_returns_tracked_change_dataclass(self, original_with_comments3_path):
        changes = extract_tracked_changes(original_with_comments3_path)
        if changes:
            assert isinstance(changes[0], TrackedChange)
