"""Tests for fuzzy matching of anchor text to target paragraphs."""

from merge_word_comments.match import (
    find_best_paragraph_match,
    find_anchor_offset,
    expand_anchor_context,
    match_comments_to_target,
)
from merge_word_comments.types import Comment, MatchResult


class TestExpandAnchorContext:
    """Expanding short anchor text with surrounding words for better matching."""

    def test_short_anchor_gets_expanded(self):
        paragraph_text = "I was walking the dog yesterday downtown"
        anchor = "the"
        expanded = expand_anchor_context(anchor, paragraph_text)
        assert len(expanded) > len(anchor)
        assert "the" in expanded

    def test_long_anchor_stays_unchanged(self):
        paragraph_text = "This is a long paragraph with many words in it that keeps going on and on and contains even more text"
        anchor = "a long paragraph with many words in it that keeps going on and on and contains even more text"
        expanded = expand_anchor_context(anchor, paragraph_text)
        assert expanded == anchor

    def test_anchor_not_in_paragraph_returns_anchor(self):
        paragraph_text = "completely different text here"
        anchor = "missing"
        expanded = expand_anchor_context(anchor, paragraph_text)
        assert expanded == anchor

    def test_anchor_at_start_of_paragraph(self):
        paragraph_text = "The dog was running in the park"
        anchor = "The"
        expanded = expand_anchor_context(anchor, paragraph_text)
        assert expanded.startswith("The")
        assert len(expanded) > 3

    def test_anchor_at_end_of_paragraph(self):
        paragraph_text = "The dog was running in the park"
        anchor = "park"
        expanded = expand_anchor_context(anchor, paragraph_text)
        assert expanded.endswith("park")
        assert len(expanded) > 4


class TestFindBestParagraphMatch:
    """Finding the best matching paragraph for a piece of anchor text."""

    def test_exact_match_returns_high_score(self):
        paragraphs = [
            "The quick brown fox jumps over the lazy dog.",
            "A completely different paragraph about cats.",
            "Yet another paragraph about birds.",
        ]
        result = find_best_paragraph_match("quick brown fox", paragraphs)
        assert result.target_paragraph_index == 0
        assert result.score >= 95

    def test_fuzzy_match_on_edited_text(self):
        paragraphs = [
            "The quick brown fox leaps gracefully over the lazy dog.",
            "A completely different paragraph about cats.",
        ]
        result = find_best_paragraph_match("quick brown fox jumps over", paragraphs)
        assert result.target_paragraph_index == 0
        assert result.score >= 60

    def test_no_match_returns_low_score(self):
        paragraphs = [
            "Something about weather.",
            "Another paragraph about cooking.",
        ]
        result = find_best_paragraph_match("quantum physics theory", paragraphs)
        assert result.score < 50

    def test_empty_paragraphs_list(self):
        result = find_best_paragraph_match("some text", [])
        assert result is None

    def test_empty_anchor_text(self):
        paragraphs = ["Some paragraph text."]
        result = find_best_paragraph_match("", paragraphs)
        assert result is not None  # should still return something (fallback)

    def test_selects_best_among_similar(self):
        paragraphs = [
            "The dog was playing in the yard with a ball.",
            "The dog was playing in the park with a frisbee.",
            "The cat was sleeping on the couch.",
        ]
        result = find_best_paragraph_match(
            "dog was playing in the park", paragraphs
        )
        assert result.target_paragraph_index == 1


class TestFindAnchorOffset:
    """Finding character-level offset of anchor text within a paragraph."""

    def test_exact_substring_found(self):
        paragraph = "The quick brown fox jumps over the lazy dog."
        offset = find_anchor_offset("brown fox", paragraph)
        assert offset == paragraph.index("brown fox")

    def test_fuzzy_substring_found(self):
        paragraph = "The quick brown fox leaps over the lazy dog."
        offset = find_anchor_offset("brown fox jumps", paragraph)
        assert offset is not None
        # Should be near the position of "brown fox"
        assert abs(offset - paragraph.index("brown fox")) < 5

    def test_no_match_returns_none(self):
        paragraph = "A completely unrelated paragraph."
        offset = find_anchor_offset("quantum physics", paragraph)
        assert offset is None


class TestMatchCommentsToTarget:
    """Integration of matching: comments against target paragraphs."""

    def test_returns_match_results(self):
        comments = [
            Comment(
                comment_id=0,
                author="Test",
                initials="T",
                date=None,
                text="A test comment",
                anchor_text="brown fox",
                anchor_context="The quick brown fox jumps",
                start_paragraph_index=0,
                end_paragraph_index=0,
                xml_element=None,
            ),
        ]
        target_paragraphs = [
            "Something unrelated.",
            "The quick brown fox jumps over the lazy dog.",
            "Another unrelated paragraph.",
        ]
        results = match_comments_to_target(comments, target_paragraphs, threshold=70)
        assert len(results) == 1
        assert isinstance(results[0], MatchResult)

    def test_matched_comment_has_correct_paragraph(self):
        comments = [
            Comment(
                comment_id=0,
                author="Test",
                initials="T",
                date=None,
                text="A comment",
                anchor_text="lazy dog",
                anchor_context="over the lazy dog",
                start_paragraph_index=0,
                end_paragraph_index=0,
                xml_element=None,
            ),
        ]
        target_paragraphs = [
            "The quick brown fox jumps over the lazy dog.",
            "A cat sat on the mat.",
        ]
        results = match_comments_to_target(comments, target_paragraphs, threshold=70)
        assert results[0].target_paragraph_index == 0

    def test_below_threshold_still_returns_with_warning(self):
        """Comments below threshold should still be included but flagged."""
        comments = [
            Comment(
                comment_id=0,
                author="Test",
                initials="T",
                date=None,
                text="A comment",
                anchor_text="quantum entanglement theory",
                anchor_context="quantum entanglement theory",
                start_paragraph_index=0,
                end_paragraph_index=0,
                xml_element=None,
            ),
        ]
        target_paragraphs = [
            "The dog was playing in the park.",
            "The cat was sleeping on the couch.",
        ]
        results = match_comments_to_target(comments, target_paragraphs, threshold=80)
        assert len(results) == 1
        assert results[0].below_threshold is True

    def test_multiple_comments_matched(self):
        comments = [
            Comment(
                comment_id=0,
                author="Test",
                initials="T",
                date=None,
                text="Comment 1",
                anchor_text="brown fox",
                anchor_context="The quick brown fox",
                start_paragraph_index=0,
                end_paragraph_index=0,
                xml_element=None,
            ),
            Comment(
                comment_id=1,
                author="Test",
                initials="T",
                date=None,
                text="Comment 2",
                anchor_text="lazy dog",
                anchor_context="the lazy dog",
                start_paragraph_index=0,
                end_paragraph_index=0,
                xml_element=None,
            ),
        ]
        target_paragraphs = [
            "The quick brown fox jumps over the lazy dog.",
        ]
        results = match_comments_to_target(comments, target_paragraphs, threshold=70)
        assert len(results) == 2
