#!/usr/bin/env python3
"""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  CSV Watcher — Automated File Monitor
  Watches a folder for new CSV/Excel files, auto-cleans
  them, and emails the results. Zero manual work.
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""

import argparse
import os
import sys
import time
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from rich.console import Console
from rich.panel import Panel
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

# Import the cleaner engine from cleaner.py
from cleaner import DataCleaner, generate_summary_report, send_email, print_summary_table

import pandas as pd

load_dotenv()
console = Console()

# Files currently being written to (avoid processing incomplete uploads)
STABLE_WAIT_SECONDS = 2
PROCESSED_MARKER = ".processed"


class CSVHandler(FileSystemEventHandler):
    """Handles new CSV/Excel files dropped into the watched folder."""

    def __init__(self, output_dir: str, no_email: bool = False):
        super().__init__()
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.no_email = no_email
        self.processing = set()

    def on_created(self, event):
        if event.is_directory:
            return
        self._handle(event.src_path)

    def on_moved(self, event):
        """Also catch files moved/renamed into the folder."""
        if event.is_directory:
            return
        self._handle(event.dest_path)

    def _handle(self, filepath: str):
        path = Path(filepath)

        # Only process CSV/Excel files
        if path.suffix.lower() not in (".csv", ".xls", ".xlsx"):
            return

        # Skip already-processed files and output files
        if path.stem.startswith("cleaned_"):
            return
        marker = path.parent / f"{path.stem}{PROCESSED_MARKER}"
        if marker.exists():
            return

        # Avoid double-processing
        if str(path) in self.processing:
            return
        self.processing.add(str(path))

        try:
            self._wait_until_stable(path)
            self._process(path)
            # Create marker so we don't re-process on restart
            marker.touch()
        except Exception as e:
            console.print(f"  [bold red]✗[/] Error processing {path.name}: {e}")
        finally:
            self.processing.discard(str(path))

    def _wait_until_stable(self, path: Path):
        """Wait until the file size stops changing (upload complete)."""
        prev_size = -1
        while True:
            curr_size = path.stat().st_size
            if curr_size == prev_size and curr_size > 0:
                break
            prev_size = curr_size
            time.sleep(STABLE_WAIT_SECONDS)

    def _process(self, path: Path):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        now_str = datetime.now().strftime("%H:%M:%S")

        console.print()
        console.print(
            f"  [bold cyan]⚡ {now_str}[/] — New file detected: "
            f"[bold white]{path.name}[/]"
        )

        # ── Load ──
        if path.suffix.lower() in (".xls", ".xlsx"):
            df = pd.read_excel(path)
        else:
            df = pd.read_csv(path)
        console.print(f"    Loaded {len(df)} rows × {len(df.columns)} columns")

        # ── Clean ──
        cleaner = DataCleaner(df)
        cleaned_df = cleaner.clean()
        summary = cleaner.summary()

        # ── Save ──
        output_path = str(self.output_dir / f"cleaned_{path.stem}_{timestamp}.xlsx")
        cleaned_df.to_excel(output_path, index=False, engine="openpyxl")

        # ── Report ──
        report_path = generate_summary_report(summary, str(path), output_path)

        # ── Display ──
        console.print(
            f"    [green]✓[/] Cleaned: {summary['original_rows']} → "
            f"{summary['cleaned_rows']} rows "
            f"([red]-{summary['rows_removed']}[/] removed)"
        )
        console.print(f"    [green]✓[/] Saved:   {output_path}")
        console.print(f"    [green]✓[/] Report:  {report_path}")

        # ── Email ──
        if not self.no_email:
            success = send_email([output_path, report_path], summary)
            if success:
                console.print(
                    f"    [green]✓[/] Emailed to "
                    f"[cyan]{os.getenv('RECIPIENT_EMAIL')}[/]"
                )
            else:
                console.print(f"    [yellow]⚠[/] Email skipped (check .env)")
        else:
            console.print(f"    [dim]Email skipped (--no-email)[/]")

        console.print(f"  [dim]{'─' * 50}[/]")


def main():
    parser = argparse.ArgumentParser(
        description="Watch a folder for new CSV files and auto-clean + email them.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python watcher.py ~/Downloads\n"
            "  python watcher.py ./inbox -o ./cleaned\n"
            "  python watcher.py ./inbox --no-email\n"
        ),
    )
    parser.add_argument(
        "watch_dir",
        nargs="?",
        default="./inbox",
        help="Folder to watch for new CSV files (default: ./inbox)",
    )
    parser.add_argument(
        "-o", "--output-dir",
        default="./cleaned",
        help="Folder to save cleaned files (default: ./cleaned)",
    )
    parser.add_argument(
        "--no-email",
        action="store_true",
        help="Skip sending email",
    )
    args = parser.parse_args()

    watch_path = Path(args.watch_dir).resolve()
    watch_path.mkdir(parents=True, exist_ok=True)
    output_path = Path(args.output_dir).resolve()

    console.print(
        Panel.fit(
            "[bold cyan]CSV Watcher[/]  [dim]— Automated Mode[/]\n"
            f"[dim]Watching:[/]  [bold]{watch_path}[/]\n"
            f"[dim]Output:[/]   [bold]{output_path}[/]\n"
            f"[dim]Email:[/]    [bold]{'OFF' if args.no_email else 'ON'}[/]",
            border_style="bright_cyan",
            padding=(1, 3),
        )
    )
    console.print("  [dim]Drop CSV/Excel files into the watched folder.[/]")
    console.print("  [dim]Press Ctrl+C to stop.[/]\n")

    handler = CSVHandler(output_dir=str(output_path), no_email=args.no_email)
    observer = Observer()
    observer.schedule(handler, str(watch_path), recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        console.print("\n  [yellow]Stopping watcher...[/]")
        observer.stop()
    observer.join()
    console.print("  [green]✓[/] Watcher stopped.\n")


if __name__ == "__main__":
    main()
