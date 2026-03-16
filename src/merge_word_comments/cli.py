"""CLI entry point for merge-word-comments."""

from pathlib import Path
from typing import Annotated

import typer
from rich.console import Console

from merge_word_comments.merge import merge_comments


app = typer.Typer(
    name="merge-word-comments",
    help="Merge comments from annotated Word documents into an updated version.",
)

console = Console()


@app.callback()
def callback() -> None:
    """Merge comments from annotated Word documents into an updated version."""


def _validate_path_exists(path: Path, label: str) -> None:
    if not path.exists():
        console.print(f"[red]Error:[/red] {label} not found: {path}")
        raise typer.Exit(code=1)
    if not path.suffix.lower() == ".docx":
        console.print(f"[red]Error:[/red] {label} must be a .docx file: {path}")
        raise typer.Exit(code=1)


@app.command()
def merge(
    updated: Annotated[
        Path,
        typer.Argument(help="Path to the updated .docx (newer version, base for output)."),
    ],
    originals: Annotated[
        list[Path],
        typer.Argument(help="One or more original .docx files containing comments/tracked changes."),
    ],
    output: Annotated[
        Path,
        typer.Option("--output", "-o", help="Output file path for the merged .docx."),
    ],
    threshold: Annotated[
        int,
        typer.Option("--threshold", "-t", help="Fuzzy match threshold 0-100."),
    ] = 80,
    verbose: Annotated[
        bool,
        typer.Option("--verbose", "-v", help="Show detailed matching info."),
    ] = False,
    intermediates: Annotated[
        bool,
        typer.Option("--intermediates", help="Save intermediate documents after each original."),
    ] = False,
    adaptive: Annotated[
        bool,
        typer.Option("--adaptive/--no-adaptive", help="Accept near-threshold matches when clearly best."),
    ] = True,
) -> None:
    """Merge comments and tracked changes from ORIGINALS into UPDATED, writing to OUTPUT."""
    _validate_path_exists(updated, "Updated file")
    for orig in originals:
        _validate_path_exists(orig, "Original file")

    merge_comments(
        updated_path=updated,
        original_paths=originals,
        output_path=output,
        threshold=threshold,
        verbose=verbose,
        intermediates=intermediates,
        adaptive=adaptive,
    )

    console.print(f"Merged output written to: {output}")
