"""Tests for the CLI interface."""

from pathlib import Path

from typer.testing import CliRunner
from docx import Document

from merge_word_comments.cli import app


runner = CliRunner()

WP_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


class TestCLIMergeCommand:
    """Tests for the 'merge' CLI command."""

    def test_merge_basic_invocation(self, updated_path,
                                     original_with_comments_path, tmp_path):
        output = tmp_path / "out.docx"
        result = runner.invoke(app, [
            "merge",
            str(updated_path),
            str(original_with_comments_path),
            "--output", str(output),
        ])
        assert result.exit_code == 0
        assert output.exists()

    def test_merge_multiple_originals(self, updated_path,
                                       original_with_comments_path,
                                       original_with_comments2_path, tmp_path):
        output = tmp_path / "out.docx"
        result = runner.invoke(app, [
            "merge",
            str(updated_path),
            str(original_with_comments_path),
            str(original_with_comments2_path),
            "--output", str(output),
        ])
        assert result.exit_code == 0
        assert output.exists()

    def test_merge_with_threshold(self, updated_path,
                                   original_with_comments_path, tmp_path):
        output = tmp_path / "out.docx"
        result = runner.invoke(app, [
            "merge",
            str(updated_path),
            str(original_with_comments_path),
            "--output", str(output),
            "--threshold", "60",
        ])
        assert result.exit_code == 0

    def test_merge_with_verbose(self, updated_path,
                                 original_with_comments_path, tmp_path):
        output = tmp_path / "out.docx"
        result = runner.invoke(app, [
            "merge",
            str(updated_path),
            str(original_with_comments_path),
            "--output", str(output),
            "--verbose",
        ])
        assert result.exit_code == 0
        # Verbose should produce more output
        assert len(result.output) > 0

    def test_merge_missing_updated_file(self, original_with_comments_path, tmp_path):
        output = tmp_path / "out.docx"
        result = runner.invoke(app, [
            "merge",
            "/nonexistent/file.docx",
            str(original_with_comments_path),
            "--output", str(output),
        ])
        assert result.exit_code != 0

    def test_merge_missing_original_file(self, updated_path, tmp_path):
        output = tmp_path / "out.docx"
        result = runner.invoke(app, [
            "merge",
            str(updated_path),
            "/nonexistent/original.docx",
            "--output", str(output),
        ])
        assert result.exit_code != 0

    def test_merge_missing_output_flag(self, updated_path,
                                        original_with_comments_path):
        result = runner.invoke(app, [
            "merge",
            str(updated_path),
            str(original_with_comments_path),
        ])
        assert result.exit_code != 0

    def test_merge_output_short_flag(self, updated_path,
                                      original_with_comments_path, tmp_path):
        output = tmp_path / "out.docx"
        result = runner.invoke(app, [
            "merge",
            str(updated_path),
            str(original_with_comments_path),
            "-o", str(output),
        ])
        assert result.exit_code == 0

    def test_merge_threshold_short_flag(self, updated_path,
                                         original_with_comments_path, tmp_path):
        output = tmp_path / "out.docx"
        result = runner.invoke(app, [
            "merge",
            str(updated_path),
            str(original_with_comments_path),
            "-o", str(output),
            "-t", "90",
        ])
        assert result.exit_code == 0

    def test_output_contains_comments(self, updated_path,
                                       original_with_comments_path, tmp_path):
        output = tmp_path / "out.docx"
        runner.invoke(app, [
            "merge",
            str(updated_path),
            str(original_with_comments_path),
            "-o", str(output),
        ])
        doc = Document(str(output))
        comments_part = doc.part._comments_part
        assert comments_part is not None
        comment_els = comments_part.element.findall(f"{{{WP_NS}}}comment")
        assert len(comment_els) >= 1
