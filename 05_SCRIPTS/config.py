from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
INBOX_DIR = PROJECT_ROOT / "00_INBOX"
JOBS_DIR = PROJECT_ROOT / "01_JOBS"
JOBS_INDEX_PATH = PROJECT_ROOT / "jobs_index.csv"
ICLOUD_DRIVE_DIR = Path.home() / "Library" / "Mobile Documents" / "com~apple~CloudDocs"
ICLOUD_WORK_DIR = ICLOUD_DRIVE_DIR / "F&T Contruction"
ICLOUD_INCOMING_DIR = ICLOUD_WORK_DIR / "HurricaneOps Incoming"
ICLOUD_PROCESSED_DIR = ICLOUD_WORK_DIR / "HurricaneOps Processed"
ICLOUD_AUTOMATION_DIR = Path.home() / "Library" / "Application Support" / "HurricaneOps"
ICLOUD_AUTO_JOBS_DIR = ICLOUD_AUTOMATION_DIR / "01_JOBS"
ICLOUD_AUTO_JOBS_INDEX_PATH = ICLOUD_AUTOMATION_DIR / "jobs_index.csv"
ICLOUD_SYNC_STATE_PATH = ICLOUD_AUTOMATION_DIR / "icloud_sync_state.json"
ICLOUD_SYNC_LOCK_PATH = ICLOUD_AUTOMATION_DIR / ".icloud_auto_ingest.lock"
ORGANIZED_PLAN_SCHEMA_VERSION = 4

JOB_SUBFOLDERS = (
    "00_Source_PDFs",
    "01_Extracted_Pages",
    "02_Schedules",
    "03_Takeoff_QA",
    "04_Quote_Workbook",
    "05_NOA_PSF_Check",
    "06_Proposal_Email",
    "07_Organized_Plan_Set",
)

ORGANIZED_PLAN_COMPARTMENTS = (
    ("floor_plans", "01_Floor_Plans", "Floor plans"),
    ("door_schedules", "02_Door_Schedules", "Door schedules"),
    ("window_schedules", "03_Window_Schedules", "Window schedules"),
    ("storefront_schedules", "04_Storefront_Schedules", "Storefront schedules"),
    ("architectural_elevations", "05_Architectural_Elevations", "Architectural elevations"),
    ("wind_pressure_elevations", "06_Wind_Pressure_Elevations", "Wind-pressure elevations"),
    ("other_plan_documents", "07_Other_Plan_Documents", "Other plan documents"),
)

EXCEL_EXTENSIONS = (".xlsx", ".xlsm", ".xltx", ".xltm", ".xls")
