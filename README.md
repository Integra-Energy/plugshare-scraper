# PlugShare Status Scraper

Small, reproducible browser-backed test scraper for finding PlugShare chargers
marked as unavailable or under repair.

It starts a normal anonymous PlugShare map session and scans a state bounding box
using the regional requests made by the public website. It does not require a
personal PlugShare account. Keep runs modest and check PlugShare's current terms
before using it beyond a test.

## Setup

```powershell
cd C:\Integra\plugshare-status-scraper
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
playwright install chromium
```

## Run

```powershell
python .\scrape_plugshare_status.py `
  --state PA `
  --output .\pennsylvania-unavailable.csv
```

To scan the continental North America bounding box and write every charger outlet:

```powershell
python .\scrape_plugshare_status.py `
  --region north-america `
  --all-statuses `
  --output .\north-america-all-chargers-by-state-province.csv
```

The CSV is sorted by the `state_or_province` column. The resolver uses the address
abbreviation first, then Canadian postal-code prefixes and US ZIP-code ranges.
Records without enough address information are grouped under `Unknown`.

The generated Excel workbook contains a summary tab and a separate sheet for each
state, province, or unresolved-address bucket.

Build the workbook from the sorted CSV:

```powershell
C:\Users\NickMatteo\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe `
  .\build_state_workbook.py `
  .\north-america-all-chargers-by-state-province.csv `
  .\north-america-all-chargers-by-state-province.xlsx
```

Create the unavailable and maintenance subset:

```powershell
python .\filter_unavailable_maintenance.py `
  .\north-america-all-chargers-by-state-province.csv `
  .\north-america-unavailable-maintenance-by-state-province.csv
```

The browser window is intentional. As of June 2, 2026, PlugShare rejects regional
map requests from Playwright's headless Chromium while allowing the same anonymous
requests in a visible Chromium window.

The default status filter includes:

- `UNDER_REPAIR`
- `OUTOFORDER`
- `UNAVAILABLE`

To scan for a specific status only, pass `--status`:

```powershell
python .\scrape_plugshare_status.py `
  --state PA `
  --status OUTOFORDER `
  --output .\pennsylvania-out-of-order.csv
```

## Coverage Notes

The current state list contains Pennsylvania (`PA`). Add another state bounding
box to `STATE_BOUNDS` in `scrape_plugshare_status.py` to scan another state.

The collector recursively subdivides busy map regions because PlugShare caps a
regional response at 250 locations. The output preserves latitude and longitude.
The current state filter uses an approximate bounding box, so chargers immediately
outside an irregular state border may appear in the CSV.
