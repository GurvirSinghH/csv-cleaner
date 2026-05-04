# 🧹 CSV Cleaner + Email Sender

Automatically clean messy CSV/Excel files and email the results.

```
Messy CSV → Clean & Standardize → Excel + Report → Email
```

## What It Cleans

| Issue | Fix |
|---|---|
| Whitespace in cells | Stripped |
| Completely empty rows | Removed |
| Duplicate rows | Removed |
| Inconsistent dates (`01/15/2023`, `March 5 2023`) | Standardized to `YYYY-MM-DD` |
| Invalid numeric values | Coerced or cleared |
| Inconsistent casing (`ENGINEERING`, `marketing`) | Title Case |
| Invalid email addresses | Cleared |

## Quick Start

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Set up email (optional)
cp .env.example .env
# Edit .env with your Gmail + App Password

# 3. Run it
python cleaner.py sample_messy.csv --preview
```

---

## 🤖 Automated Mode (Watcher)

**Don't want to run anything manually?** Start the watcher and just drop files into a folder.

```bash
# Start watching the ./inbox folder (auto-created)
python watcher.py

# Watch a specific folder (e.g. Downloads)
python watcher.py ~/Downloads

# Custom input + output folders
python watcher.py ./inbox -o ./cleaned

# Watch without email
python watcher.py --no-email
```

**How it works:**
1. The watcher monitors a folder for new `.csv` / `.xlsx` files
2. When a file appears, it waits for the upload to finish
3. Auto-cleans the data and saves a clean `.xlsx` + report
4. Emails both files to you (if configured)
5. Marks the file as processed (won't re-process on restart)

> Drop a file → get a clean Excel in your inbox. That's it.

### Run in the background

```bash
# Run as a background process
nohup python watcher.py ~/Downloads -o ./cleaned &> watcher.log &

# Check the log
tail -f watcher.log

# Stop it
kill $(pgrep -f "python watcher.py")
```

---

## Manual Mode

```bash
# Basic — clean and email
python cleaner.py data.csv

# Custom output path
python cleaner.py data.csv -o my_clean_data.xlsx

# Skip email, just clean
python cleaner.py data.csv --no-email

# Show before/after preview
python cleaner.py data.csv --preview

# All options
python cleaner.py data.csv -o output.xlsx --preview --no-email
```

## Gmail Setup

1. Go to [Google Account Security](https://myaccount.google.com/security)
2. Enable **2-Step Verification**
3. Go to **App Passwords** → Generate one for "Mail"
4. Paste the 16-char password in your `.env` file

## Output Files

Each run produces:
- `cleaned_<name>_<timestamp>.xlsx` — The cleaned dataset
- `cleaned_<name>_<timestamp>_report.txt` — Summary of all changes made
