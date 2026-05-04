#!/usr/bin/env python3
"""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  CSV Cleaner + Email Sender
  Cleans messy CSV/Excel files and emails the results.
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""

import argparse
import os
import re
import smtplib
import sys
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.progress import Progress, SpinnerColumn, TextColumn

# ──────────────────────────────────────────────
#  Setup
# ──────────────────────────────────────────────

console = Console()
load_dotenv()

TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")


# ──────────────────────────────────────────────
#  Data Cleaning Engine
# ──────────────────────────────────────────────

class DataCleaner:
    """Cleans a pandas DataFrame by removing junk, fixing formats, and standardizing values."""

    def __init__(self, df: pd.DataFrame):
        self.original = df.copy()
        self.df = df.copy()
        self.log: list[str] = []

    def _record(self, action: str, count: int):
        """Log a cleaning action with its impact count."""
        if count > 0:
            self.log.append(f"{action}: {count}")

    # ── Step 1: Strip whitespace from all string cells & column names ──
    def strip_whitespace(self):
        self.df.columns = [col.strip() for col in self.df.columns]
        str_cols = self.df.select_dtypes(include=["object"]).columns
        for col in str_cols:
            self.df[col] = self.df[col].astype(str).str.strip()
            # Turn whitespace-only strings back into NaN
            self.df[col] = self.df[col].replace(r"^\s*$", pd.NA, regex=True)
            self.df[col] = self.df[col].replace("nan", pd.NA)
        self._record("Whitespace stripped in cells", len(str_cols))
        return self

    # ── Step 2: Drop completely empty rows ──
    def drop_empty_rows(self):
        before = len(self.df)
        self.df.dropna(how="all", inplace=True)
        self.df.reset_index(drop=True, inplace=True)
        removed = before - len(self.df)
        self._record("Empty rows removed", removed)
        return self

    # ── Step 3: Remove duplicate rows ──
    def drop_duplicates(self):
        before = len(self.df)
        self.df.drop_duplicates(inplace=True)
        self.df.reset_index(drop=True, inplace=True)
        removed = before - len(self.df)
        self._record("Duplicate rows removed", removed)
        return self

    # ── Step 4: Standardize date columns ──
    def standardize_dates(self):
        fixed = 0
        for col in self.df.columns:
            if any(kw in col.lower() for kw in ["date", "time", "created", "joined", "updated"]):
                original_vals = self.df[col].copy()
                self.df[col] = pd.to_datetime(self.df[col], errors="coerce", dayfirst=False)
                changed = (original_vals.notna() & self.df[col].notna()).sum()
                coerced_to_nat = (original_vals.notna() & self.df[col].isna()).sum()
                fixed += int(changed)
                if coerced_to_nat > 0:
                    self._record(f"Invalid dates set to blank in '{col}'", int(coerced_to_nat))
        self._record("Date values standardized", fixed)
        return self

    # ── Step 5: Clean numeric columns ──
    def clean_numerics(self):
        fixed = 0
        for col in self.df.columns:
            if any(kw in col.lower() for kw in ["salary", "amount", "price", "cost", "revenue", "qty", "quantity"]):
                self.df[col] = pd.to_numeric(self.df[col], errors="coerce")
                fixed += 1
        self._record("Numeric columns coerced", fixed)
        return self

    # ── Step 6: Standardize text casing ──
    def standardize_casing(self):
        fixed = 0
        for col in self.df.columns:
            if any(kw in col.lower() for kw in ["department", "category", "status", "type", "role"]):
                self.df[col] = self.df[col].astype(str).str.title()
                self.df[col] = self.df[col].replace("Nan", pd.NA)
                fixed += 1
        self._record("Text columns title-cased", fixed)
        return self

    # ── Step 7: Validate emails ──
    def validate_emails(self):
        email_pattern = re.compile(r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$")
        flagged = 0
        for col in self.df.columns:
            if "email" in col.lower():
                mask = self.df[col].apply(
                    lambda x: bool(email_pattern.match(str(x))) if pd.notna(x) else True
                )
                flagged += int((~mask).sum())
                self.df.loc[~mask, col] = pd.NA
        self._record("Invalid emails cleared", flagged)
        return self

    # ── Run the full pipeline ──
    def clean(self) -> pd.DataFrame:
        """Execute all cleaning steps in order."""
        self.strip_whitespace()
        self.drop_empty_rows()
        self.drop_duplicates()
        self.standardize_dates()
        self.clean_numerics()
        self.standardize_casing()
        self.validate_emails()
        return self.df

    # ── Summary statistics ──
    def summary(self) -> dict:
        """Return a summary comparing original vs cleaned data."""
        return {
            "original_rows": len(self.original),
            "cleaned_rows": len(self.df),
            "rows_removed": len(self.original) - len(self.df),
            "original_columns": len(self.original.columns),
            "remaining_nulls": int(self.df.isna().sum().sum()),
            "actions": self.log,
        }


# ──────────────────────────────────────────────
#  Report Generator
# ──────────────────────────────────────────────

def generate_summary_report(summary: dict, input_file: str, output_file: str) -> str:
    """Generate a plain-text summary report and save it to disk."""
    report_path = output_file.replace(".xlsx", "_report.txt")

    lines = [
        "=" * 60,
        "  DATA CLEANING SUMMARY REPORT",
        f"  Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "=" * 60,
        "",
        f"  Input File  : {input_file}",
        f"  Output File : {output_file}",
        "",
        "─" * 60,
        "  METRICS",
        "─" * 60,
        f"  Original Rows    : {summary['original_rows']}",
        f"  Cleaned Rows     : {summary['cleaned_rows']}",
        f"  Rows Removed     : {summary['rows_removed']}",
        f"  Remaining Nulls  : {summary['remaining_nulls']}",
        "",
        "─" * 60,
        "  CLEANING ACTIONS PERFORMED",
        "─" * 60,
    ]
    for action in summary["actions"]:
        lines.append(f"  ✓ {action}")

    lines.append("")
    lines.append("=" * 60)
    report_text = "\n".join(lines)

    Path(report_path).write_text(report_text)
    return report_path


# ──────────────────────────────────────────────
#  Email Sender
# ──────────────────────────────────────────────

def send_email(attachments: list[str], summary: dict):
    """Send the cleaned file + report as email attachments."""
    smtp_server = os.getenv("SMTP_SERVER", "smtp.gmail.com")
    smtp_port = int(os.getenv("SMTP_PORT", 587))
    sender = os.getenv("SENDER_EMAIL")
    password = os.getenv("SENDER_PASSWORD")
    recipient = os.getenv("RECIPIENT_EMAIL", sender)

    if not sender or not password:
        console.print(
            "[bold red]✗[/] Email credentials not found. "
            "Copy .env.example → .env and fill in your details.",
        )
        return False

    # Build message
    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = f"📊 Cleaned Data Report — {datetime.now().strftime('%b %d, %Y')}"

    body = (
        f"Hi,\n\n"
        f"Your cleaned dataset is attached.\n\n"
        f"Quick stats:\n"
        f"  • Original rows : {summary['original_rows']}\n"
        f"  • Clean rows    : {summary['cleaned_rows']}\n"
        f"  • Rows removed  : {summary['rows_removed']}\n"
        f"  • Remaining nulls: {summary['remaining_nulls']}\n\n"
        f"See the attached report for full details.\n\n"
        f"— CSV Cleaner Bot 🤖"
    )
    msg.attach(MIMEText(body, "plain"))

    # Attach files
    for filepath in attachments:
        with open(filepath, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename={Path(filepath).name}",
        )
        msg.attach(part)

    # Send
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(sender, password)
            server.send_message(msg)
        return True
    except smtplib.SMTPAuthenticationError:
        console.print(
            "[bold red]✗[/] Authentication failed. "
            "Check your email/password in .env (use a Gmail App Password).",
        )
        return False
    except Exception as e:
        console.print(f"[bold red]✗[/] Email failed: {e}")
        return False


# ──────────────────────────────────────────────
#  Rich Console UI
# ──────────────────────────────────────────────

def print_banner():
    console.print(
        Panel.fit(
            "[bold cyan]CSV Cleaner + Email Sender[/]\n"
            "[dim]Clean messy data → Excel → Email[/]",
            border_style="bright_cyan",
            padding=(1, 4),
        )
    )


def print_summary_table(summary: dict):
    table = Table(
        title="🧹 Cleaning Results",
        show_header=True,
        header_style="bold magenta",
        border_style="bright_cyan",
        padding=(0, 2),
    )
    table.add_column("Metric", style="cyan", min_width=20)
    table.add_column("Value", style="green", justify="right", min_width=10)

    table.add_row("Original Rows", str(summary["original_rows"]))
    table.add_row("Cleaned Rows", str(summary["cleaned_rows"]))
    table.add_row("Rows Removed", str(summary["rows_removed"]))
    table.add_row("Remaining Nulls", str(summary["remaining_nulls"]))
    console.print()
    console.print(table)

    if summary["actions"]:
        console.print()
        console.print("[bold]Actions performed:[/]")
        for action in summary["actions"]:
            console.print(f"  [green]✓[/] {action}")


def print_data_preview(df: pd.DataFrame, title: str):
    table = Table(title=title, border_style="dim", padding=(0, 1), show_lines=True)
    for col in df.columns:
        table.add_column(str(col), style="cyan", overflow="fold", max_width=25)
    for _, row in df.head(5).iterrows():
        table.add_row(*[str(v) if pd.notna(v) else "[dim]—[/]" for v in row])
    console.print()
    console.print(table)


# ──────────────────────────────────────────────
#  Main
# ──────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Clean CSV/Excel files and optionally email the results.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python cleaner.py sample_messy.csv\n"
            "  python cleaner.py data.csv -o clean_output.xlsx\n"
            "  python cleaner.py data.csv --no-email\n"
        ),
    )
    parser.add_argument("input", help="Path to the messy CSV or Excel file")
    parser.add_argument("-o", "--output", help="Output Excel file path (auto-generated if omitted)")
    parser.add_argument("--no-email", action="store_true", help="Skip sending email")
    parser.add_argument("--preview", action="store_true", help="Show before/after data preview")
    args = parser.parse_args()

    print_banner()

    # ── Load ──
    input_path = Path(args.input)
    if not input_path.exists():
        console.print(f"[bold red]✗[/] File not found: {input_path}")
        sys.exit(1)

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        console=console,
    ) as progress:
        # Read input
        task = progress.add_task("Reading input file...", total=None)
        if input_path.suffix.lower() in (".xls", ".xlsx"):
            df = pd.read_excel(input_path)
        else:
            df = pd.read_csv(input_path)
        progress.update(task, description="[green]✓[/] File loaded", completed=True)

        if args.preview:
            print_data_preview(df, "📥 Raw Input (first 5 rows)")

        # Clean
        task = progress.add_task("Cleaning data...", total=None)
        cleaner = DataCleaner(df)
        cleaned_df = cleaner.clean()
        summary = cleaner.summary()
        progress.update(task, description="[green]✓[/] Data cleaned", completed=True)

        # Save
        output_path = args.output or str(
            input_path.parent / f"cleaned_{input_path.stem}_{TIMESTAMP}.xlsx"
        )
        task = progress.add_task("Saving Excel file...", total=None)
        cleaned_df.to_excel(output_path, index=False, engine="openpyxl")
        progress.update(task, description="[green]✓[/] Excel saved", completed=True)

        # Report
        task = progress.add_task("Generating report...", total=None)
        report_path = generate_summary_report(summary, str(input_path), output_path)
        progress.update(task, description="[green]✓[/] Report generated", completed=True)

    # ── Display results ──
    print_summary_table(summary)

    if args.preview:
        print_data_preview(cleaned_df, "📤 Cleaned Output (first 5 rows)")

    console.print()
    console.print(f"  [bold]📁 Excel:[/]  [link=file://{Path(output_path).resolve()}]{output_path}[/]")
    console.print(f"  [bold]📄 Report:[/] [link=file://{Path(report_path).resolve()}]{report_path}[/]")

    # ── Email ──
    if not args.no_email:
        console.print()
        with Progress(
            SpinnerColumn(),
            TextColumn("[progress.description]{task.description}"),
            console=console,
        ) as progress:
            task = progress.add_task("Sending email...", total=None)
            success = send_email([output_path, report_path], summary)
            if success:
                progress.update(task, description="[green]✓[/] Email sent!", completed=True)
                console.print(
                    f"  [bold green]✓[/] Emailed to [cyan]{os.getenv('RECIPIENT_EMAIL')}[/]"
                )
            else:
                progress.update(task, description="[yellow]⚠[/] Email skipped", completed=True)
    else:
        console.print("\n  [dim]Email skipped (--no-email)[/]")

    console.print()


if __name__ == "__main__":
    main()
