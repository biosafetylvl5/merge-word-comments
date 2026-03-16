"""Data types for merge-word-comments."""

from dataclasses import dataclass, field
from typing import Optional

from lxml.etree import _Element


@dataclass(frozen=True)
class Comment:
    """A comment extracted from a Word document."""

    comment_id: int
    author: str
    initials: Optional[str]
    date: Optional[str]
    text: str
    anchor_text: str
    anchor_context: str
    start_paragraph_index: int
    end_paragraph_index: int
    xml_element: Optional[_Element]
    source_heading: str = ""  # enclosing heading keyword from source document


@dataclass(frozen=True)
class TrackedChange:
    """A tracked change (insertion or deletion) from a Word document."""

    change_type: str  # "insert" or "delete"
    author: str
    date: Optional[str]
    content: str
    paragraph_context: str  # original (pre-change) paragraph text
    paragraph_index: int
    char_offset: Optional[int] = None
    context_paragraph_offset: int = 0  # 0 = same para, -1 = preceding, +1 = following, etc.
    paragraph_context_current: str = ""  # current (with changes) paragraph text, for fallback matching
    local_context: str = ""  # ~100 chars around the change offset for sub-paragraph matching
    source_heading: str = ""  # enclosing heading keyword from source document
    xml_elements: list[_Element] = field(default_factory=list)


@dataclass(frozen=True)
class MatchResult:
    """Result of matching a comment to a target paragraph."""

    comment: Comment
    target_paragraph_index: int
    score: float
    anchor_offset: Optional[int]
    below_threshold: bool = False
    target_end_paragraph_index: Optional[int] = None

    def effective_end_index(self) -> int:
        """Return the end paragraph index, defaulting to start if single-para."""
        if self.target_end_paragraph_index is not None:
            return self.target_end_paragraph_index
        return self.target_paragraph_index


@dataclass(frozen=True)
class FailureRecord:
    """A comment or tracked change that could not be merged."""

    kind: str  # "comment" or "tracked_change"
    source_file: str
    author: str
    date: Optional[str]
    content_preview: str
    reason: str
    source_paragraph_index: int
    char_offset: Optional[int] = None
    anchor_text: str = ""
    anchor_context: str = ""
    change_type: Optional[str] = None  # "insert" or "delete" for tracked changes
    best_match_score: Optional[float] = None
    best_match_paragraph_index: Optional[int] = None
    best_match_paragraph_preview: Optional[str] = None
    threshold: int = 80
