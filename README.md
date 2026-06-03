# HurricaneOps

HurricaneOps starts each bid job from one PDF plan-set folder and one Excel quote
workbook template placed directly in `00_INBOX`. The plan-set folder may contain
one combined multi-page PDF or multiple PDF sheets.

## Setup

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
```

## Run The Intake Pipeline

1. Add exactly one folder containing plan-set PDFs and exactly one Excel template
   directly inside `00_INBOX`.
2. Run:

```bash
.venv/bin/python 05_SCRIPTS/run_bid_pipeline.py
```

The script creates a dated folder in `01_JOBS`, copies the PDF plan set and quote
template, extracts text-backed schedule candidates, writes QA and PSF/NOA review
reports, creates a formula-preserving quote workbook draft, and writes a proposal
email draft.

Each new job also includes `07_Organized_Plan_Set` with flattened review copies:

```text
00_Indexed_Sheets
01_Floor_Plans
02_Door_Schedules
03_Window_Schedules
04_Storefront_Schedules
05_Architectural_Elevations
06_Wind_Pressure_Elevations
07_Other_Plan_Documents
00_Download_Packages
source_file_manifest.csv
sheet_index.csv
```

The original plan-set hierarchy remains preserved in `00_Source_PDFs`. The
organizer splits combined PDFs into one-page indexed sheets, reads extractable
cover-sheet indexes and title blocks, and uses page text as a fallback. It does
not assume a fixed sheet identifier such as `A-701`. A sheet can be copied into
more than one review compartment when appropriate.
The download packages include separate ZIP files for floor plans, architectural
elevations, wind-pressure elevations, door schedules, window schedules,
storefront schedules, other plan documents, and the complete organized PDF set.
The dashboard also exposes each extracted schedule CSV separately, including a
mixed `glazing_schedule_candidates.csv` file when a drawing places window and
exterior-door openings together in a glazing schedule. For quote review, those
mixed glazing rows stay together in drawing order on the generated workbook's
`Storefronts` worksheet as the project-specific storefront glazing package.
Validated architectural and garage-door schedule rows are written in drawing
order on the generated workbook's `Doors` worksheet. Each run writes
`door_workbook_transfer_audit.json` and aborts if the saved workbook does not
reconcile row-for-row with the validated door data.
Window-only rows from a mixed glazing schedule are also copied to the generated
workbook's `Windows` worksheet without removing them from the unified glazing
review package. Each run writes `window_workbook_transfer_audit.json` and aborts
if the saved window columns do not reconcile with the validated window rows.
The Storefronts worksheet is also reconciled after save through
`storefront_workbook_transfer_audit.json`. Explicit schedule-level glazing
requirements such as U-factor and SHGC are preserved in
`glazing_thermal_requirements.json` and copied into the review workbook.
Window and storefront `Glass TYPE` is carried from the schedule `Brand/Product`
column when the drawings do not provide a separate glass-type column. Material
and finish stay blank unless explicitly extracted, and HurricaneOps adds QA
review items for missing material/finish fields. Door `FRAME FINISH` is carried
from the second value in combined `DOOR / FRAME` material cells such as
`PAINTED WOOD, PAINTED WOOD`.
The dashboard's Critical Issue Action Center links each critical item to its
best matching copied PDF, CSV, workbook, or QA file. Estimators can record an
accept/deny/RFI decision or enter a verified value that is written directly into
the filled quote workbook. Workbook Count Check tables are regenerated for
Windows, Storefronts, and Doors so each extracted mark or door number is counted
against the generated unit columns.
When an
earlier completed job is selected, the dashboard offers a one-click action to
create or refresh the organized downloads from its preserved source PDFs.

The automation uses PDF text and table extraction first. A modular image
fallback can render weak PDF pages at high resolution, preprocess the image,
detect schedule regions and table grids, and OCR individual table cells when
OpenCV, pytesseract, and the local Tesseract executable are available. It does
not invent missing quantities or PSF values, and it does not overwrite existing
workbook formulas.

`opencv-python-headless` and `pytesseract` install from `requirements.txt`.
Image-only OCR also requires a system `tesseract` executable on the Mac. When it
is absent, HurricaneOps records an explicit audit warning and does not guess
values.

Each pipeline job includes `extracted_schedules.json`, `extraction_audit.json`,
and `uncertain_schedule_rows.csv`. Only validated structured rows with
confidence of at least `0.85` are eligible for quote-workbook population.

To run the extractor separately:

```bash
python extract_schedules.py input.pdf --project-address "160 NW 44 St" --output ./output
```

## Run The Local App

```bash
.venv/bin/streamlit run app.py
```

The local Streamlit app stages one combined plan-set PDF or multiple PDF sheets
with an Excel quote template, runs the same command-line pipeline, displays
estimator QA metrics, blocks quote issue when CRITICAL items exist, provides the
review package downloads, and tracks dashboard jobs in `jobs_index.csv`.

Estimator QA resolution state is saved locally inside each job at
`03_Takeoff_QA/estimator_review_state.json`. The dashboard also renders
first-page previews for the copied source sheets. Downloads remain available as
review artifacts; unresolved CRITICAL items prevent the dashboard from treating
the quote as final.

When unresolved CRITICAL issues remain, the Estimator Risk Acceptance section
can record an estimator-level decision in `approval_decision.json`. This unlocks
the package for estimator review only. The dashboard keeps the CRITICAL warnings
and labels the ZIP `Review Package - Not Final`; it does not mark the quote final.

## iCloud Automatic Intake

While the local dashboard is open, HurricaneOps checks this dedicated iCloud
Drive queue every two minutes:

```text
iCloud Drive/F&T Contruction/HurricaneOps Incoming
```

Place each new bid inside its own project folder. Add one combined plan-set PDF
or the separate plan-set PDF sheets to that folder. The automation uses the
standard workbook from the iCloud `Quote Template` folder unless the project
folder contains one Excel workbook.

The project folder name becomes the project name by default. To provide a cleaner
name, address, or notes, add an optional `project.json` file:

```json
{
  "project_name": "Example Project",
  "address": "123 Main Street, Miami FL",
  "notes": "Bid due Friday."
}
```

Completed ZIP packages and pipeline summaries are copied to:

```text
iCloud Drive/F&T Contruction/HurricaneOps Processed
```

The processed iCloud result also includes the browsable `07_Organized_Plan_Set`
folder beside the ZIP package. The existing iCloud `Qoutes` folder is not scanned
or modified. You can also run an immediate queue check from the Streamlit
dashboard.

The dashboard uses a small runtime copy under
`~/Library/Application Support/HurricaneOps` for automatically generated jobs.
This keeps automatic intake separate from manually staged jobs and works around
macOS blocking background LaunchAgents from reading iCloud Drive. Reinstall or
refresh that runtime with:

```bash
05_SCRIPTS/install_icloud_automation.sh
```
