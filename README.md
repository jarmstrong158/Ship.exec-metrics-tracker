# ShipExec Metrics

Automated warehouse shipping report generator for [ShipExec](https://thinclient.shipexec.com). Scrapes History and Detailed Report tables via Selenium, then builds per-region Excel summaries with picked/packed counts, boxcounts, and staff breakdowns.

## Business Impact

Built for a warehouse operation processing 200+ daily orders across multiple regions and a team of 6 staff. Replaced a fully manual reporting process that required a manager to log into ShipExec, manually cross-reference two separate report tables, and build Excel summaries by hand each day.

What this tool eliminated:

- Manual daily data pulls from two separate ShipExec report views
- Hand-calculation of per-staff picked/packed counts and boxcounts across 4 regional teams
- The weekend/Monday reporting gap — a ShipExec limitation that previously caused 3 days of shipment data to collapse into a single undifferentiated report

**Result:** Saves 4 staff members 1–1.5 hours of manual reporting work per day. At current compensation levels, that recovers approximately 1,300 labor hours and $35,000–$60,000 in annual labor value — from a tool with zero licensing cost.

**Stack:** Python, Selenium, BeautifulSoup, pandas — no RPA platform license required.

## What It Does

Each warehouse region (e.g. NE, SE, CTR, WC) has staff identified by initials in the Consignee Reference field. The script:

1. Launches Chrome and navigates to ShipExec (you log in manually)
2. Scrapes the History table (tracking numbers, shipper/consignee references)
3. Scrapes the Detailed Report (tracking numbers, emails for region assignment)
4. Merges the two datasets on tracking number
5. Counts picked/packed per staff member per region and tracks boxcounts
6. Outputs an Excel workbook with a `History` sheet and a `RegionSummary` sheet

### Weekend/Monday Problem

ShipExec groups Saturday, Sunday, and Monday shipments under the same date range. Without intervention, you'd get three days of data lumped together with no way to tell which day a shipment belongs to.

This script solves that by running each day separately and subtracting prior days:

1. **Saturday** -- scrapes Saturday's shipments and saves them as the weekend baseline
2. **Sunday** -- scrapes Saturday+Sunday together, then subtracts Saturday's data to isolate Sunday-only
3. **Monday** -- scrapes the full Sat-Mon window, then subtracts both Saturday and Sunday to isolate Monday-only

**Tuesday-Friday** don't have this problem -- each day has its own date range in ShipExec, so the script scrapes and reports directly.

## Requirements

- Python 3.10+
- Google Chrome installed
- A ShipExec account with access to History and Detailed Report

## Setup

```bash
git clone <this-repo>
cd shipexec-metrics
pip install -r requirements.txt
```

## Usage

```bash
python metrics.py
```

On first run, you'll be guided through setup:

1. **Region configuration** -- accept the defaults (NE, SE, CTR, WC) or define your own regions with custom email patterns
2. **Sort preference** -- choose alphabetical by last initial (default) or custom ordering per region
3. A `metrics_config.json` file is saved alongside the script and reused on future runs

Then select a report type from the menu:

```
Select the type of report to run:
  1 = Monday
  2 = Tuesday-Friday
  3 = Saturday
  4 = Sunday
  5 = Manage Regions
  6 = Exit
```

Chrome will launch automatically. **Log in to ShipExec in the browser window**, then press Enter in the terminal to continue. The script handles the rest.

## Configuration

All configuration is stored in `metrics_config.json` (auto-generated, git-ignored). You can edit it through the **Manage Regions** menu option (choice 5), which lets you:

- Add or remove regions
- Edit the display order of staff initials per region
- Switch between alphabetical and custom sort modes

### How Initials Are Discovered

Staff initials are **not hardcoded**. On each run the script discovers them from the scraped data:

- Initials that appear **2+ times** or in **both packer and picker positions** are treated as confirmed staff and added to the main summary
- Initials that appear only once in a single position are placed in a **Review** section at the bottom of the summary with context clues (role, frequency, and similar-initial suggestions)
- New initials are automatically detected and you're prompted to confirm their position in the sort order
- Configured initials that don't appear in a given day's data still show up with 0 counts

## Output Files

Reports are saved to `~/Documents/`:

| File | Written By |
|---|---|
| `history_with_email_and_summary_automated.xlsx` | Tuesday-Friday, Saturday |
| `history_with_email_and_summary_SUN.xlsx` | Sunday |
| `history_with_email_and_summary_SUN_raw.xlsx` | Sunday (full scrape before isolation) |
| `history_with_email_and_summary_MON.xlsx` | Monday |
| `history.xlsx` | Every run (intermediate History scrape) |

Each workbook contains:

- **History** sheet -- tracking numbers, shipper references, consignee references, emails, and region assignments
- **RegionSummary** sheet -- per-region columns with staff initials, picked counts, packed counts, totals, boxcounts, and (if applicable) a review section for unrecognized initials

## Project Structure

```
shipexec-metrics/
  metrics.py            # The entire script
  requirements.txt      # Python dependencies
  metrics_config.json   # Auto-generated on first run (git-ignored)
  .gitignore
  README.md
```

## Building a Standalone Executable

If you want to share this with someone who doesn't have Python installed, you can package it as a standalone `.exe`:

```bash
pip install pyinstaller
pyinstaller --onefile metrics.py
```

The executable will be in the `dist/` folder. The user will still need Google Chrome installed and a ShipExec account, but they won't need Python or any dependencies.

## Security

- **No credentials are stored in the code.** The browser launches and you log in manually.
- `metrics_config.json` contains only region names, email pattern prefixes (e.g. `ne@`), and display preferences. It is git-ignored by default.

## License

MIT
