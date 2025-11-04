import os
from datetime import datetime, date
from typing import List, Optional

import pandas as pd
import requests
from dotenv import load_dotenv
from flask import (
    Flask,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from flask_caching import Cache
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import func
from werkzeug.utils import secure_filename
from functools import wraps

# Load environment variables
load_dotenv()

# Directory configuration
BASE_DIR = os.getcwd()
DATA_DIR = os.path.join(BASE_DIR, "data")
LOGO_DIR = os.path.join(BASE_DIR, "static", "logos")
PROJECT_IMAGE_DIR = os.path.join(BASE_DIR, "static", "project_images")
MAIN_DIR = os.path.join(BASE_DIR, "documents")  # Added missing MAIN_DIR definition

# Ensure directories exist
for directory in [DATA_DIR, LOGO_DIR, PROJECT_IMAGE_DIR, MAIN_DIR]:
    os.makedirs(directory, exist_ok=True)

# File paths
EMPLOYEE_FILE = os.path.join(DATA_DIR, "employees.csv")
COMPANY_FILE = os.path.join(DATA_DIR, "companies.csv")
PROJECT_FILE = os.path.join(DATA_DIR, "projects.csv")
HARM_FILE = os.path.join(DATA_DIR, "harm_drive_vehicles.csv")
WAHOO_FILE = os.path.join(DATA_DIR, "wahoo_pool_vehicles.csv")
LOCKYER_FILE = os.path.join(DATA_DIR, "lockyer_sheds_vehicles.csv")
EXCEL_FILE = os.path.join(DATA_DIR, "companies.xlsx")  # Added missing EXCEL_FILE
PROJECTS_FILE = os.path.join(DATA_DIR, "projects.xlsx")  # Added missing PROJECTS_FILE
HARM_DRIVE_FILE = os.path.join(DATA_DIR, "HarmDriveData.xlsx")  # Added missing HARM_DRIVE_FILE
WAHOO_VEHICLES_FILE = os.path.join(DATA_DIR, "wahoo_pool_vehicles.xlsx")  # Added missing WAHOO_VEHICLES_FILE

VEHICLE_COLUMNS = [
    "Plate",
    "Type",
    "VIN",
    "Rego Renewal Date",
    "Insurance Renewal (CTP) Date",
    "Value",
    "Transfer Fee",
]

FLEET_CONFIG = {
    "wahoo": {
        "file": WAHOO_FILE,
        "title": "Wahoo Pool Vehicles",
        "heading": "Wahoo Pool Construction Vehicles",
    },
    "lockyer": {
        "file": LOCKYER_FILE,
        "title": "Lockyer Sheds Vehicles",
        "heading": "Lockyer Sheds Fleet",
    },
    "harm": {
        "file": HARM_FILE,
        "title": "Harm Drive Vehicles",
        "heading": "Harm Drive Fleet",
    },
}

# Flask App configuration
app = Flask(__name__, template_folder="Templates")
app.secret_key = os.getenv("FLASK_SECRET_KEY", "fallback_secret_key")
cache = Cache(app, config={"CACHE_TYPE": "simple"})

# Database configuration
default_sqlite_path = os.path.join(DATA_DIR, "mkpd_cms.db")
app.config["SQLALCHEMY_DATABASE_URI"] = os.getenv(
    "DATABASE_URL",
    f"sqlite:///{default_sqlite_path}",
)
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
db = SQLAlchemy(app)

# Admin password
PASSWORD = os.getenv("ADMIN_PASSWORD", "Admin.123.")

# File configuration
ALLOWED_EXTENSIONS = {"pdf", "docx", "jpg", "jpeg", "png"}

# Utility functions
def allowed_file(filename):
    """Check if a file has an allowed extension"""
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

def calculate_expiry(rego_date):
    """
    Calculate days until registration expiry
    :param rego_date: Registration renewal date as string 'YYYY-MM-DD'
    :return: Number of days until expiry or None if invalid
    """
    try:
        today = datetime.today()
        if isinstance(rego_date, datetime):
            target_date = rego_date
        elif isinstance(rego_date, date):
            target_date = datetime.combine(rego_date, datetime.min.time())
        else:
            target_date = datetime.strptime(str(rego_date), '%Y-%m-%d')
        return (target_date - today).days
    except Exception:
        return None

def normalise_date(value):
    """Normalise a value into YYYY-MM-DD string format or empty string."""
    if value in (None, ""):
        return ""

    try:
        parsed = pd.to_datetime(value)
        if pd.isna(parsed):
            return ""
        return parsed.strftime("%Y-%m-%d")
    except Exception:
        # Fallback to string representation (e.g., already formatted)
        value_str = str(value).strip()
        try:
            datetime.strptime(value_str, "%Y-%m-%d")
            return value_str
        except Exception:
            return ""

def coerce_numeric(value):
    """Safely convert values to floats for storage."""
    try:
        if value in ("", None):
            return None
        return float(value)
    except (TypeError, ValueError):
        return None


def parse_date(value):
    """Convert arbitrary date-like values into a date object."""
    normalized = normalise_date(value)
    if not normalized:
        return None
    try:
        return datetime.strptime(normalized, "%Y-%m-%d").date()
    except ValueError:
        return None


class Vehicle(db.Model):
    """Database representation of a fleet vehicle."""

    __tablename__ = "vehicles"

    id = db.Column(db.Integer, primary_key=True)
    plate = db.Column(db.String(32), unique=True, nullable=False)
    type = db.Column(db.String(120), nullable=False)
    vin = db.Column(db.String(64), nullable=False)
    rego_renewal_date = db.Column(db.Date, nullable=True)
    insurance_renewal_date = db.Column(db.Date, nullable=True)
    value = db.Column(db.Float, nullable=True)
    transfer_fee = db.Column(db.Float, nullable=True)
    fleet = db.Column(db.String(32), nullable=False, index=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        """Serialize the model into the structure expected by templates."""

        rego_date_str = (
            self.rego_renewal_date.strftime("%Y-%m-%d")
            if self.rego_renewal_date
            else ""
        )
        insurance_date_str = (
            self.insurance_renewal_date.strftime("%Y-%m-%d")
            if self.insurance_renewal_date
            else ""
        )

        return {
            "Plate": self.plate,
            "Type": self.type,
            "VIN": self.vin,
            "Rego Renewal Date": rego_date_str,
            "Insurance Renewal (CTP) Date": insurance_date_str,
            "Value": self.value,
            "Transfer Fee": self.transfer_fee,
            "Expiry": calculate_expiry(rego_date_str) if rego_date_str else None,
        }


class Employee(db.Model):
    """Employee directory records."""

    __tablename__ = "employees"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False)
    family_name = db.Column(db.String(120), nullable=False)
    tfn = db.Column(db.String(20), nullable=True)
    abn = db.Column(db.String(20), nullable=True)
    address = db.Column(db.String(255), nullable=True)
    email = db.Column(db.String(255), nullable=True)
    phone = db.Column(db.String(64), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    __table_args__ = (
        db.UniqueConstraint("name", "family_name", name="uq_employee_name"),
    )


class Company(db.Model):
    """Company registry records."""

    __tablename__ = "companies"

    id = db.Column(db.Integer, primary_key=True)
    company_name = db.Column(db.String(255), nullable=False)
    registration_date = db.Column(db.Date, nullable=True)
    abn = db.Column(db.String(32), nullable=False, unique=True)
    acn = db.Column(db.String(32), nullable=True)
    company_type = db.Column(db.String(120), nullable=True)
    registered_address = db.Column(db.String(255), nullable=True)
    qbcc_license_number = db.Column(db.String(64), nullable=True)
    documents_folder = db.Column(db.String(255), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class Project(db.Model):
    """Project portfolio records."""

    __tablename__ = "projects"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(255), nullable=False, unique=True)
    description = db.Column(db.Text, nullable=False)
    start_date = db.Column(db.Date, nullable=True)
    end_date = db.Column(db.Date, nullable=True)
    image_filename = db.Column(db.String(255), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


def load_vehicle_data(file_path, fleet_slug=None):
    """Load vehicle records using the database, with CSV fallback."""
    if fleet_slug:
        try:
            vehicles = (
                Vehicle.query.filter_by(fleet=fleet_slug)
                .order_by(Vehicle.rego_renewal_date.asc(), Vehicle.plate.asc())
                .all()
            )
            if vehicles:
                return [vehicle.to_dict() for vehicle in vehicles]
        except Exception as exc:
            print(f"Failed to load vehicles for {fleet_slug} from database: {exc}")

    if not os.path.exists(file_path) or os.path.getsize(file_path) == 0:
        return []

    try:
        df = pd.read_csv(file_path)
    except pd.errors.EmptyDataError:
        return []
    except Exception as exc:
        print(f"Failed to read vehicle data from {file_path}: {exc}")
        return []

    for column in VEHICLE_COLUMNS:
        if column not in df.columns:
            df[column] = pd.NA

    df = df[VEHICLE_COLUMNS].copy()
    df["Rego Renewal Date"] = df["Rego Renewal Date"].apply(normalise_date)
    df["Insurance Renewal (CTP) Date"] = df["Insurance Renewal (CTP) Date"].apply(normalise_date)
    df["Value"] = pd.to_numeric(df["Value"], errors="coerce")
    df["Transfer Fee"] = pd.to_numeric(df["Transfer Fee"], errors="coerce")
    df["Expiry"] = df["Rego Renewal Date"].apply(lambda value: calculate_expiry(value) if value else None)

    df.sort_values(by=["Rego Renewal Date", "Plate"], inplace=True, na_position="last")

    return df.to_dict(orient="records")

def format_date_ddmmyyyy(date):
    """Format a date object to DD/MM/YYYY string format"""
    if pd.notnull(date):
        return datetime.strptime(str(date), "%Y-%m-%d").strftime("%d/%m/%Y")
    return None

# Template utility functions
@app.template_filter('datetimeformat')
def datetimeformat(value):
    """Template filter to format dates"""
    try:
        return datetime.strptime(value, '%Y-%m-%d').strftime('%d/%m/%Y')
    except Exception:
        return value

@app.template_global()
def get_company_logo_static(company_name, abn):
    """Return static URL for company logo or default image"""
    logo_filename = f"{company_name.replace(' ', '_')}_{abn}.png"
    logo_path = os.path.join(LOGO_DIR, logo_filename)
    
    if os.path.exists(logo_path):
        return url_for("static", filename=f"logos/{logo_filename}")
    else:
        return url_for("static", filename="logos/default_logo.png")

# Authentication middleware
def login_required(f):
    """Decorator to protect routes requiring authentication"""
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("logged_in"):
            flash("Please log in first.", "warning")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated

# Setup functions
def setup_default_files():
    """Create default data files if they don't exist"""
    # Create default projects file
    if not os.path.exists(PROJECTS_FILE):
        pd.DataFrame(columns=["Project Name", "Description", "Start Date", "End Date", "Image Filename"]).to_excel(
            PROJECTS_FILE, index=False, engine="openpyxl")
        print(f"✅ Created default projects file: {PROJECTS_FILE}")

    # Create default employees file
    if not os.path.exists(EMPLOYEE_FILE):
        pd.DataFrame(columns=["Name", "Family Name", "TFN", "ABN", "Address", "Email", "Phone"]).to_csv(
            EMPLOYEE_FILE, index=False)
        print(f"✅ Created default employees file: {EMPLOYEE_FILE}")

    # Create default companies file
    if not os.path.exists(EXCEL_FILE):
        pd.DataFrame(columns=["Company Name", "Registration Date", "ABN", "ACN", "Type", "Registered Address",
                             "QBCC License Number", "Documents"]).to_excel(
            EXCEL_FILE, index=False, engine="openpyxl")
        print(f"✅ Created default companies file: {EXCEL_FILE}")

    # Setup Harm Drive data file
    setup_harm_drive_file()

    # Setup vehicle files for each fleet
    setup_vehicle_file(
        WAHOO_FILE,
        [
            {
                "Plate": "371RLI",
                "Type": "Mitsubishi Triton",
                "VIN": "MMAENKA40BD006265",
                "Rego Renewal Date": "2025-03-25",
                "Insurance Renewal (CTP) Date": "2025-05-01",
                "Value": 48000,
                "Transfer Fee": 305.00,
            },
            {
                "Plate": "295RMK",
                "Type": "Mazda BT-50",
                "VIN": "MM0UNY0W400891903",
                "Rego Renewal Date": "2025-03-06",
                "Insurance Renewal (CTP) Date": "2025-04-15",
                "Value": 52500,
                "Transfer Fee": 315.00,
            },
            {
                "Plate": "975SMU",
                "Type": "Mitsubishi Triton",
                "VIN": "MMAJNKB40CD018749",
                "Rego Renewal Date": "2025-03-20",
                "Insurance Renewal (CTP) Date": "2025-04-25",
                "Value": 49200,
                "Transfer Fee": 305.00,
            },
        ],
        "wahoo",
    )

    setup_vehicle_file(
        LOCKYER_FILE,
        [
            {
                "Plate": "452GHT",
                "Type": "Isuzu NPR75",
                "VIN": "JAANPR75L87123456",
                "Rego Renewal Date": "2025-02-11",
                "Insurance Renewal (CTP) Date": "2025-03-30",
                "Value": 68500,
                "Transfer Fee": 420.00,
            },
            {
                "Plate": "918KLS",
                "Type": "Hino 300",
                "VIN": "JHDFS8JJ70K005432",
                "Rego Renewal Date": "2025-04-05",
                "Insurance Renewal (CTP) Date": "2025-05-18",
                "Value": 61250,
                "Transfer Fee": 395.00,
            },
            {
                "Plate": "625PLQ",
                "Type": "Fuso Canter",
                "VIN": "JLDCEKJ02MF004512",
                "Rego Renewal Date": "2025-01-28",
                "Insurance Renewal (CTP) Date": "2025-03-10",
                "Value": 64890,
                "Transfer Fee": 405.00,
            },
        ],
        "lockyer",
    )

    setup_vehicle_file(
        HARM_FILE,
        [
            {
                "Plate": "HDX001",
                "Type": "Ford Ranger",
                "VIN": "MNAUMFF80GW123456",
                "Rego Renewal Date": "2025-06-12",
                "Insurance Renewal (CTP) Date": "2025-07-01",
                "Value": 49800,
                "Transfer Fee": 312.50,
            },
            {
                "Plate": "HDX274",
                "Type": "Toyota HiAce",
                "VIN": "JTFRA3AP508765432",
                "Rego Renewal Date": "2025-08-03",
                "Insurance Renewal (CTP) Date": "2025-09-15",
                "Value": 43250,
                "Transfer Fee": 284.10,
            },
            {
                "Plate": "HDX518",
                "Type": "Isuzu D-MAX",
                "VIN": "MPATFS40JFT018765",
                "Rego Renewal Date": "2025-05-19",
                "Insurance Renewal (CTP) Date": "2025-06-30",
                "Value": 45675,
                "Transfer Fee": 297.80,
            },
        ],
        "harm",
    )

    seed_employees_from_csv()
    seed_companies_from_excel()
    seed_projects_from_disk()
    export_employees_to_csv()
    export_companies_to_excel()
    export_projects_to_disk()

def setup_harm_drive_file():
    """Create HarmDriveData.xlsx if it doesn't exist"""
    if not os.path.exists(HARM_DRIVE_FILE):
        data = {
            "Plate": ["HDX001", "HDX274", "HDX518"],
            "Type": ["Ford Ranger", "Toyota HiAce", "Isuzu D-MAX"],
            "VIN": ["MNAUMFF80GW123456", "JTFRA3AP508765432", "MPATFS40JFT018765"],
            "Rego Renewal Date": ["2025-06-12", "2025-08-03", "2025-05-19"],
            "Insurance Renewal (CTP) Date": ["2025-07-01", "2025-09-15", "2025-06-30"],
            "Expiry": [None, None, None],
            "Value": [49800, 43250, 45675],
            "Transfer Fee": [312.50, 284.10, 297.80]
        }
        df = pd.DataFrame(data)
        df.to_excel(HARM_DRIVE_FILE, index=False, engine="openpyxl")
        print(f"✅ Created HarmDriveData.xlsx file")


def sync_vehicle_csv(fleet_slug: str) -> None:
    """Write the current fleet records to the CSV on disk."""

    config = FLEET_CONFIG.get(fleet_slug)
    if not config:
        return

    records = (
        Vehicle.query.filter_by(fleet=fleet_slug)
        .order_by(Vehicle.rego_renewal_date.asc(), Vehicle.plate.asc())
        .all()
    )

    rows: List[dict] = []
    for vehicle in records:
        rows.append(
            {
                "Plate": vehicle.plate,
                "Type": vehicle.type,
                "VIN": vehicle.vin,
                "Rego Renewal Date": vehicle.rego_renewal_date.strftime("%Y-%m-%d")
                if vehicle.rego_renewal_date
                else "",
                "Insurance Renewal (CTP) Date": vehicle.insurance_renewal_date.strftime("%Y-%m-%d")
                if vehicle.insurance_renewal_date
                else "",
                "Value": vehicle.value,
                "Transfer Fee": vehicle.transfer_fee,
            }
        )

    df = pd.DataFrame(rows, columns=VEHICLE_COLUMNS)
    df.to_csv(config["file"], index=False)


def seed_vehicle_records(fleet_slug, file_path, default_rows):
    """Populate the database for a fleet when records are missing."""
    if db is None:
        return

    existing_plates = {
        vehicle.plate
        for vehicle in Vehicle.query.filter_by(fleet=fleet_slug).all()
    }

    rows_to_seed = default_rows

    if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
        try:
            df = pd.read_csv(file_path)
            if not df.empty:
                rows_to_seed = df.to_dict(orient="records")
        except pd.errors.EmptyDataError:
            rows_to_seed = default_rows
        except Exception as exc:
            print(f"Failed to load CSV data for {fleet_slug}: {exc}")

    added = 0
    for row in rows_to_seed:
        plate = str(row.get("Plate", "")).strip()
        if not plate or plate in existing_plates:
            continue

        vehicle = Vehicle(
            plate=plate,
            type=str(row.get("Type", "")).strip() or "Unknown",
            vin=str(row.get("VIN", "")).strip() or "Unknown",
            rego_renewal_date=parse_date(row.get("Rego Renewal Date")),
            insurance_renewal_date=parse_date(row.get("Insurance Renewal (CTP) Date")),
            value=coerce_numeric(row.get("Value")),
            transfer_fee=coerce_numeric(row.get("Transfer Fee")),
            fleet=fleet_slug,
        )
        db.session.add(vehicle)
        existing_plates.add(plate)
        added += 1

    if added:
        db.session.commit()
        print(f"✅ Seeded {added} {fleet_slug} vehicle records into the database")
        sync_vehicle_csv(fleet_slug)


def setup_vehicle_file(file_path, default_rows, fleet_slug):
    """Ensure a vehicle CSV exists and database records are seeded."""
    if not os.path.exists(file_path) or os.path.getsize(file_path) == 0:
        df = pd.DataFrame(default_rows, columns=VEHICLE_COLUMNS)
        df.to_csv(file_path, index=False)
        print(f"✅ Created default vehicle file: {file_path}")

    seed_vehicle_records(fleet_slug, file_path, default_rows)


def seed_employees_from_csv() -> None:
    """Populate the employees table from the legacy CSV file if needed."""

    if Employee.query.count():
        return

    columns = ["Name", "Family Name", "TFN", "ABN", "Address", "Email", "Phone"]
    if not os.path.exists(EMPLOYEE_FILE):
        pd.DataFrame(columns=columns).to_csv(EMPLOYEE_FILE, index=False)
        return

    try:
        df = pd.read_csv(EMPLOYEE_FILE)
    except (pd.errors.EmptyDataError, FileNotFoundError):
        return

    if df.empty:
        return

    for row in df.itertuples(index=False):
        employee = Employee(
            name=str(getattr(row, "Name", "")).strip(),
            family_name=str(getattr(row, "Family_Name", getattr(row, "Family Name", ""))).strip(),
            tfn=str(getattr(row, "TFN", "")).strip() or None,
            abn=str(getattr(row, "ABN", "")).strip() or None,
            address=str(getattr(row, "Address", "")).strip() or None,
            email=str(getattr(row, "Email", "")).strip() or None,
            phone=str(getattr(row, "Phone", "")).strip() or None,
        )

        if not employee.name or not employee.family_name:
            continue

        existing = Employee.query.filter_by(
            name=employee.name,
            family_name=employee.family_name,
        ).first()
        if existing:
            continue

        db.session.add(employee)

    db.session.commit()


def export_employees_to_csv() -> None:
    """Persist the employees table back to the CSV file."""

    records = Employee.query.order_by(Employee.name.asc(), Employee.family_name.asc()).all()
    rows = [
        {
            "Name": employee.name,
            "Family Name": employee.family_name,
            "TFN": employee.tfn or "",
            "ABN": employee.abn or "",
            "Address": employee.address or "",
            "Email": employee.email or "",
            "Phone": employee.phone or "",
        }
        for employee in records
    ]

    df = pd.DataFrame(rows, columns=["Name", "Family Name", "TFN", "ABN", "Address", "Email", "Phone"])
    df.to_csv(EMPLOYEE_FILE, index=False)


def seed_companies_from_excel() -> None:
    """Populate the companies table from the Excel workbook."""

    if Company.query.count():
        return

    columns = [
        "Company Name",
        "Registration Date",
        "ABN",
        "ACN",
        "Type",
        "Registered Address",
        "QBCC License Number",
        "Documents",
    ]

    if not os.path.exists(EXCEL_FILE):
        pd.DataFrame(columns=columns).to_excel(EXCEL_FILE, index=False, engine="openpyxl")
        return

    try:
        df = pd.read_excel(EXCEL_FILE, engine="openpyxl")
    except FileNotFoundError:
        return

    if df.empty:
        return

    for row in df.to_dict(orient="records"):
        company_name = str(row.get("Company Name", "")).strip()
        abn = str(row.get("ABN", "")).strip()
        if not company_name or not abn:
            continue

        registration_date = parse_date(row.get("Registration Date"))

        company = Company(
            company_name=company_name,
            registration_date=registration_date,
            abn=abn,
            acn=str(row.get("ACN", "")).strip() or None,
            company_type=str(row.get("Type", "")).strip() or None,
            registered_address=str(row.get("Registered Address", "")).strip() or None,
            qbcc_license_number=str(row.get("QBCC License Number", "")).strip() or None,
            documents_folder=str(row.get("Documents", "")).strip() or None,
        )

        if Company.query.filter_by(abn=company.abn).first():
            continue

        db.session.add(company)

    db.session.commit()


def export_companies_to_excel() -> None:
    """Persist the companies table back to the Excel workbook."""

    records = Company.query.order_by(Company.company_name.asc()).all()
    rows = []
    for company in records:
        rows.append(
            {
                "Company Name": company.company_name,
                "Registration Date": company.registration_date.strftime("%Y-%m-%d")
                if company.registration_date
                else "",
                "ABN": company.abn,
                "ACN": company.acn or "",
                "Type": company.company_type or "",
                "Registered Address": company.registered_address or "",
                "QBCC License Number": company.qbcc_license_number or "",
                "Documents": company.documents_folder or "",
            }
        )

    df = pd.DataFrame(rows, columns=[
        "Company Name",
        "Registration Date",
        "ABN",
        "ACN",
        "Type",
        "Registered Address",
        "QBCC License Number",
        "Documents",
    ])
    df.to_excel(EXCEL_FILE, index=False, engine="openpyxl")


def seed_projects_from_disk() -> None:
    """Seed projects from the existing CSV or Excel file if the table is empty."""

    if Project.query.count():
        return

    dataframe: Optional[pd.DataFrame] = None

    if os.path.exists(PROJECTS_FILE):
        try:
            dataframe = pd.read_excel(PROJECTS_FILE, engine="openpyxl")
        except Exception:
            dataframe = None

    if dataframe is None and os.path.exists(PROJECT_FILE):
        try:
            dataframe = pd.read_csv(PROJECT_FILE)
        except Exception:
            dataframe = None

    if dataframe is None or dataframe.empty:
        return

    for row in dataframe.to_dict(orient="records"):
        name = str(row.get("Project Name", "")).strip()
        if not name:
            continue

        project = Project(
            name=name,
            description=str(row.get("Description", "")).strip() or "No description provided.",
            start_date=parse_date(row.get("Start Date")),
            end_date=parse_date(row.get("End Date")),
            image_filename=str(row.get("Image Filename", "")).strip() or None,
        )

        if Project.query.filter_by(name=project.name).first():
            continue

        db.session.add(project)

    db.session.commit()


def export_projects_to_disk() -> None:
    """Persist the projects table to both CSV and Excel for compatibility."""

    records = (
        Project.query.order_by(
            Project.start_date.is_(None),
            Project.start_date.asc(),
            Project.name.asc(),
        ).all()
    )
    rows = []
    for project in records:
        rows.append(
            {
                "Project Name": project.name,
                "Description": project.description,
                "Start Date": project.start_date.strftime("%Y-%m-%d") if project.start_date else "",
                "End Date": project.end_date.strftime("%Y-%m-%d") if project.end_date else "",
                "Image Filename": project.image_filename or "",
            }
        )

    df = pd.DataFrame(rows, columns=[
        "Project Name",
        "Description",
        "Start Date",
        "End Date",
        "Image Filename",
    ])
    df.to_excel(PROJECTS_FILE, index=False, engine="openpyxl")
    df.to_csv(PROJECT_FILE, index=False)

# API functions
@cache.cached(timeout=600)  # Cache for 10 minutes
def get_qld_construction_news():
    """Fetch recent Queensland construction news from News API"""
    api_key = os.getenv('NEWS_API_KEY', 'bfb864ec86be43f49b257cb04ff2ab0f')
    endpoint = 'https://newsapi.org/v2/everything'
    
    params = {
        'q': 'Queensland construction',
        'language': 'en',
        'sortBy': 'publishedAt',
        'pageSize': 20,
        'apiKey': api_key
    }

    try:
        response = requests.get(endpoint, params=params)
        response.raise_for_status()
        
        data = response.json()
        articles = data.get('articles', [])
        
        filtered_articles = []
        for article in articles:
            filtered_articles.append({
                'title': article.get('title', 'No title'),
                'url': article.get('url', '#'),
                'publishedAt': article.get('publishedAt', 'Unknown date')
            })

        return filtered_articles

    except requests.exceptions.RequestException as e:
        print(f"Error fetching news: {e}")
        return []
    except ValueError as e:
        print(f"Error parsing response JSON: {e}")
        return []

# Routes
@app.route("/login", methods=["GET", "POST"])
def login():
    """Handle user login"""
    if request.method == "POST":
        password = request.form.get("password")
        if password == PASSWORD:
            session["logged_in"] = True
            flash("Login successful!", "success")
            return redirect(url_for("index"))
        else:
            flash("Incorrect password. Try again.", "error")
    return render_template("login.html", current_year=datetime.utcnow().year)

@app.route("/logout")
def logout():
    """Handle user logout"""
    session.clear()
    flash("Logged out", "info")
    return redirect(url_for("login"))

@app.route("/")
@login_required
def index():
    """Main dashboard page"""
    try:
        # Load vehicle data
        wahoo_vehicles = load_vehicle_data(WAHOO_FILE, "wahoo")
        lockyer_vehicles = load_vehicle_data(LOCKYER_FILE, "lockyer")
        harm_drive_vehicles = load_vehicle_data(HARM_FILE, "harm")
        vehicles = wahoo_vehicles + lockyer_vehicles + harm_drive_vehicles

        # Get construction news
        news_items = get_qld_construction_news()

        # Pagination for news
        PER_PAGE = 5
        page = request.args.get('page', 1, type=int)
        total_pages = (len(news_items) + PER_PAGE - 1) // PER_PAGE
        page = max(1, min(page, total_pages))
        paginated_news = news_items[(page - 1) * PER_PAGE: page * PER_PAGE]

        return render_template(
            "index.html",
            vehicles=vehicles,
            wahoo_vehicles=wahoo_vehicles,
            lockyer_vehicles=lockyer_vehicles,
            news_items=paginated_news,
            page=page,
            total_pages=total_pages,
        )
    except Exception as e:
        flash(f"Error loading data: {e}", "error")
        return render_template("index.html", vehicles=[], news_items=[], page=1, total_pages=1)

@app.route("/navbar")
def navbar():
    """Serve the navbar template"""
    return render_template("navbar.html")

@app.route("/employees")
@login_required
def employees():
    """Display list of employees"""
    try:
        employees_list = (
            Employee.query.order_by(Employee.name.asc(), Employee.family_name.asc()).all()
        )
        return render_template("employees.html", employees=employees_list)
    except Exception as e:
        app.logger.exception("Error loading employees")
        flash(f"Error loading employees: {e}", "danger")
        return redirect(url_for("index"))

@app.route("/search_employee", methods=["GET", "POST"])
@login_required
def search_employee():
    """Search for an employee by name and family name"""
    selected_employees: List[Employee] = []

    if request.method == "POST":
        search_name = request.form.get("search_name", "").strip()
        search_family_name = request.form.get("search_family_name", "").strip()

        if not search_name or not search_family_name:
            flash("Both Name and Family Name are required!", "warning")
            return render_template("search_employee.html", selected_employees=[])

        try:
            selected_employees = (
                Employee.query.filter(
                    func.lower(Employee.name) == search_name.lower(),
                    func.lower(Employee.family_name) == search_family_name.lower(),
                )
                .order_by(Employee.created_at.desc())
                .all()
            )

            if not selected_employees:
                flash("No employees found with the specified criteria.", "info")
        except Exception as e:
            app.logger.exception("Error searching employees")
            flash(f"Error searching employees: {e}", "error")

    return render_template("search_employee.html", selected_employees=selected_employees)

@app.route("/add_employee", methods=["GET", "POST"])
@login_required
def add_employee():
    """Add a new employee"""
    if request.method == "POST":
        try:
            name = request.form["name"].strip()
            family_name = request.form["family_name"].strip()

            if not name or not family_name:
                flash("Name and Family Name are required.", "warning")
                return redirect(url_for("add_employee"))

            existing = Employee.query.filter(
                func.lower(Employee.name) == name.lower(),
                func.lower(Employee.family_name) == family_name.lower(),
            ).first()
            if existing:
                flash("This employee already exists in the directory.", "info")
                return redirect(url_for("employees"))

            employee = Employee(
                name=name,
                family_name=family_name,
                tfn=request.form.get("tfn", "").strip() or None,
                abn=request.form.get("abn", "").strip() or None,
                address=request.form.get("address", "").strip() or None,
                email=request.form.get("email", "").strip() or None,
                phone=request.form.get("phone", "").strip() or None,
            )
            db.session.add(employee)
            db.session.commit()

            export_employees_to_csv()

            flash("Employee added successfully!", "success")
            return redirect(url_for("employees"))
        except Exception as e:
            db.session.rollback()
            app.logger.exception("Error saving employee")
            flash(f"Error saving employee: {e}", "danger")
            return redirect(url_for("add_employee"))

    return render_template("add_employee.html")

@app.route("/companies")
@login_required
def companies():
    """Display list of companies"""
    try:
        companies_list = Company.query.order_by(Company.company_name.asc()).all()
        return render_template("companies.html", companies=companies_list)
    except Exception as e:
        app.logger.exception("Error loading companies")
        flash(f"Error loading companies: {e}", "danger")
        return redirect(url_for("index"))

@app.route("/add_company", methods=["GET", "POST"])
@login_required
def add_company():
    """Add a new company"""
    try:
        if request.method == "POST":
            # Collect data from the form
            company_name = request.form["company_name"].strip()
            registration_date = request.form["registration_date"].strip()
            abn = request.form["abn"].strip()
            acn = request.form.get("acn", "").strip()
            company_type = request.form.get("type", "").strip()
            registered_address = request.form.get("registered_address", "").strip()
            qbcc_license_number = request.form.get("qbcc_license_number", "").strip()

            # Validate required fields
            if not company_name or not abn:
                flash("Company Name and ABN are required!", "error")
                return redirect(url_for("add_company"))

            # Validate ABN format
            import re
            if not re.match(r'^\d{11}$', abn):
                flash("Invalid ABN format. It must be 11 digits.", "error")
                return redirect(url_for("add_company"))

            # Ensure ABN uniqueness in the database
            existing = Company.query.filter(func.lower(Company.abn) == abn.lower()).first()
            if existing:
                flash("A company with this ABN already exists!", "error")
                return redirect(url_for("add_company"))

            # Create a unique folder for the company
            folder_name = secure_filename(f"{company_name}_{abn}")
            folder_path = os.path.join(MAIN_DIR, folder_name)
            os.makedirs(folder_path, exist_ok=True)

            # Handle mandatory ASIC Extract upload
            asic_extract = request.files.get("asic_extract")
            if not asic_extract or not allowed_file(asic_extract.filename):
                flash("ASIC Extract is required and must be a valid file!", "error")
                return redirect(url_for("add_company"))

            # Save the ASIC Extract file
            asic_extract.save(os.path.join(folder_path, secure_filename(asic_extract.filename)))

            # Handle optional document uploads
            for field_name in ["company_registration", "logo"]:
                file = request.files.get(field_name)
                if file and file.filename and allowed_file(file.filename):
                    file.save(os.path.join(folder_path, secure_filename(file.filename)))

            registration_date_obj = parse_date(registration_date)

            company = Company(
                company_name=company_name,
                registration_date=registration_date_obj,
                abn=abn,
                acn=acn or None,
                company_type=company_type or None,
                registered_address=registered_address or None,
                qbcc_license_number=qbcc_license_number or None,
                documents_folder=folder_name,
            )

            db.session.add(company)
            db.session.commit()

            export_companies_to_excel()

            flash("Company added successfully!", "success")
            return redirect(url_for("view_company"))

        # Render the Add Company form
        return render_template("add_company.html", title="Add Company")

    except Exception as e:
        db.session.rollback()
        app.logger.exception("Error adding company")
        flash(f"Error adding company: {e}", "error")
        return redirect(url_for("view_company"))

@app.route("/view_company", methods=["GET", "POST"])
@login_required
def view_company():
    """View company details"""
    try:
        companies = Company.query.order_by(Company.company_name.asc()).all()
        if not companies:
            flash("No companies found. Please add a company first.", "info")
            return redirect(url_for("add_company"))

        if request.method == "POST":
            company_id = request.form.get("company_id")
            if not company_id:
                flash("Please select a company.", "warning")
                return redirect(url_for("view_company"))

            try:
                company_id_int = int(company_id)
            except (TypeError, ValueError):
                flash("Invalid company selection.", "error")
                return redirect(url_for("view_company"))

            company = Company.query.filter_by(id=company_id_int).first()
            if not company:
                flash("Company not found.", "error")
                return redirect(url_for("view_company"))

            folder_name = company.documents_folder
            folder_path = os.path.join(MAIN_DIR, folder_name)
            if not os.path.exists(folder_path):
                flash("Document folder does not exist.", "error")
                return redirect(url_for("view_company"))

            documents = [
                doc
                for doc in os.listdir(folder_path)
                if os.path.isfile(os.path.join(folder_path, doc))
            ]

            return render_template(
                "company_details.html",
                title="Company Details",
                company=company,
                documents=documents,
            )

        return render_template("view_company.html", title="View Companies", companies=companies)

    except Exception as e:
        app.logger.exception("Error loading companies")
        flash(f"Error loading companies: {e}", "error")
        return redirect(url_for("index"))

@app.route("/projects")
@login_required
def projects():
    """Display list of projects"""
    try:
        projects_list = (
            Project.query.order_by(
                Project.start_date.is_(None),
                Project.start_date.asc(),
                Project.name.asc(),
            ).all()
        )
        return render_template("projects.html", projects=projects_list)
    except Exception as e:
        flash(f"Error loading projects: {e}", "danger")
        return redirect(url_for("index"))

@app.route("/add_project", methods=["GET", "POST"])
@login_required
def add_project():
    """Add a new project"""
    if request.method == "POST":
        try:
            # Extract form data
            project_name = request.form.get("project_name", "").strip()
            project_description = request.form.get("project_description", "").strip()
            start_date = request.form.get("start_date", "").strip()
            end_date = request.form.get("end_date", "").strip()

            # Validate mandatory fields
            if not project_name or not project_description or not start_date or not end_date:
                flash("All fields except the image are required!", "error")
                return redirect(url_for("add_project"))

            existing = Project.query.filter(func.lower(Project.name) == project_name.lower()).first()
            if existing:
                flash("A project with this name already exists.", "warning")
                return redirect(url_for("view_projects"))

            # Handle image upload
            image = request.files.get("project_image")
            if image and allowed_file(image.filename):
                filename = secure_filename(image.filename)
                image.save(os.path.join(PROJECT_IMAGE_DIR, filename))
            else:
                filename = None  # No image uploaded or invalid file

            project = Project(
                name=project_name,
                description=project_description,
                start_date=parse_date(start_date),
                end_date=parse_date(end_date),
                image_filename=filename,
            )

            db.session.add(project)
            db.session.commit()

            export_projects_to_disk()

            flash("Project added successfully!", "success")
            return redirect(url_for("view_projects"))
        except Exception as e:
            db.session.rollback()
            app.logger.exception("Error adding project")
            flash(f"Error adding project: {str(e)}", "error")
            return redirect(url_for("add_project"))

    return render_template("add_project.html")

@app.route("/view_projects")
@login_required
def view_projects():
    """View all projects"""
    try:
        projects = (
            Project.query.order_by(
                Project.start_date.is_(None),
                Project.start_date.asc(),
                Project.name.asc(),
            ).all()
        )
        if not projects:
            flash("No projects found. Please add a project first.", "info")
            return redirect(url_for("add_project"))

        return render_template("view_projects.html", projects=projects)
    except Exception as e:
        app.logger.exception("Error loading projects")
        flash(f"Error loading projects: {str(e)}", "error")
        return redirect(url_for("index"))

@app.route("/view_project/<project_name>")
@login_required
def view_project(project_name):
    """View details of a specific project"""
    try:
        project = Project.query.filter(func.lower(Project.name) == project_name.lower()).first()
        if not project:
            flash(f"Project '{project_name}' not found.", "error")
            return redirect(url_for("view_projects"))

        return render_template("view_project.html", project=project)
    except Exception as e:
        app.logger.exception("Error retrieving project")
        flash(f"Error retrieving project: {str(e)}", "error")
        return redirect(url_for("view_projects"))

@app.route("/download/<company_folder>/<filename>")
@login_required
def download_file(company_folder, filename):
    """Download a company document"""
    company_folder = secure_filename(company_folder)
    filename = secure_filename(filename)

    filepath = os.path.abspath(os.path.join(MAIN_DIR, company_folder, filename))

    # Prevent path traversal attacks
    if not filepath.startswith(os.path.abspath(MAIN_DIR)):
        flash("Unauthorized access detected!", "error")
        return redirect(url_for("view_company"))

    if not os.path.exists(filepath):
        flash("File not found!", "error")
        return redirect(url_for("view_company"))

    return send_file(filepath, as_attachment=True)

def render_vehicles_page(fleet_slug):
    """Render the vehicles page for a given fleet."""
    config = FLEET_CONFIG.get(fleet_slug)
    if not config:
        flash("Unknown vehicle fleet selected.", "danger")
        return redirect(url_for("index"))

    vehicles = load_vehicle_data(config["file"], fleet_slug)
    return render_template(
        "vehicles.html",
        page_title=config["title"],
        heading=config["heading"],
        vehicles=vehicles,
        add_action=url_for("add_vehicle_record", fleet_slug=fleet_slug),
        update_action=url_for("update_vehicle_record", fleet_slug=fleet_slug),
        delete_action=url_for("delete_vehicle_record", fleet_slug=fleet_slug),
    )


def _redirect_to_fleet(fleet_slug):
    """Redirect to the correct vehicles page for a fleet."""
    if fleet_slug == "wahoo":
        return redirect(url_for("wahoo_vehicles"))
    if fleet_slug == "lockyer":
        return redirect(url_for("lockyer_vehicles"))
    if fleet_slug == "harm":
        return redirect(url_for("harm_drive"))
    return redirect(url_for("vehicles_page_route", fleet_slug=fleet_slug))


@app.route("/vehicles/<fleet_slug>")
@login_required
def vehicles_page_route(fleet_slug):
    return render_vehicles_page(fleet_slug)


@app.route("/wahoo_vehicles")
@login_required
def wahoo_vehicles():
    return render_vehicles_page("wahoo")


@app.route("/lockyer_vehicles")
@login_required
def lockyer_vehicles():
    return render_vehicles_page("lockyer")


@app.route("/harm_drive")
@login_required
def harm_drive():
    return render_vehicles_page("harm")


@app.route("/vehicles/<fleet_slug>/add", methods=["POST"])
@login_required
def add_vehicle_record(fleet_slug):
    if fleet_slug not in FLEET_CONFIG:
        flash("Unknown vehicle fleet selected.", "danger")
        return redirect(url_for("index"))
    try:
        plate = request.form.get("plate", "").strip().upper()
        vehicle_type = request.form.get("type", "").strip()
        vin = request.form.get("vin", "").strip().upper()

        if not plate or not vehicle_type or not vin:
            flash("Plate, Type, and VIN are required to add a vehicle.", "warning")
            return _redirect_to_fleet(fleet_slug)

        if Vehicle.query.filter_by(plate=plate).first():
            flash("A vehicle with this plate already exists.", "info")
            return _redirect_to_fleet(fleet_slug)

        rego_date = request.form.get("rego_renewal_date") or request.form.get("expiry_date")
        insurance_date = request.form.get("insurance_renewal_date")

        vehicle = Vehicle(
            plate=plate,
            type=vehicle_type,
            vin=vin,
            rego_renewal_date=parse_date(rego_date),
            insurance_renewal_date=parse_date(insurance_date),
            value=coerce_numeric(request.form.get("value")),
            transfer_fee=coerce_numeric(request.form.get("transfer_fee")),
            fleet=fleet_slug,
        )

        db.session.add(vehicle)
        db.session.commit()
        sync_vehicle_csv(fleet_slug)

        flash("Vehicle added successfully!", "success")
    except Exception as exc:
        db.session.rollback()
        app.logger.exception("Error adding vehicle")
        flash(f"Unable to add vehicle: {exc}", "error")

    return _redirect_to_fleet(fleet_slug)


@app.route("/vehicles/<fleet_slug>/update", methods=["POST"])
@login_required
def update_vehicle_record(fleet_slug):
    if fleet_slug not in FLEET_CONFIG:
        flash("Unknown vehicle fleet selected.", "danger")
        return redirect(url_for("index"))
    try:
        plate = request.form.get("plate_to_update", "").strip().upper()
        if not plate:
            flash("No vehicle selected for update.", "warning")
            return _redirect_to_fleet(fleet_slug)

        vehicle = Vehicle.query.filter_by(plate=plate, fleet=fleet_slug).first()
        if not vehicle:
            flash("Vehicle not found.", "error")
            return _redirect_to_fleet(fleet_slug)

        new_rego_date = parse_date(request.form.get("new_rego_renewal_date"))
        if new_rego_date is None:
            flash("Please provide a valid registration renewal date.", "warning")
            return _redirect_to_fleet(fleet_slug)

        vehicle.rego_renewal_date = new_rego_date
        db.session.commit()
        sync_vehicle_csv(fleet_slug)

        flash("Vehicle updated successfully!", "success")
    except Exception as exc:
        db.session.rollback()
        app.logger.exception("Error updating vehicle")
        flash(f"Unable to update vehicle: {exc}", "error")

    return _redirect_to_fleet(fleet_slug)


@app.route("/vehicles/<fleet_slug>/delete", methods=["POST"])
@login_required
def delete_vehicle_record(fleet_slug):
    if fleet_slug not in FLEET_CONFIG:
        flash("Unknown vehicle fleet selected.", "danger")
        return redirect(url_for("index"))
    try:
        plate = request.form.get("plate_to_delete", "").strip().upper()
        if not plate:
            flash("No vehicle selected for deletion.", "warning")
            return _redirect_to_fleet(fleet_slug)

        vehicle = Vehicle.query.filter_by(plate=plate, fleet=fleet_slug).first()
        if not vehicle:
            flash("Vehicle not found.", "error")
            return _redirect_to_fleet(fleet_slug)

        db.session.delete(vehicle)
        db.session.commit()
        sync_vehicle_csv(fleet_slug)

        flash("Vehicle removed successfully!", "success")
    except Exception as exc:
        db.session.rollback()
        app.logger.exception("Error deleting vehicle")
        flash(f"Unable to delete vehicle: {exc}", "error")

    return _redirect_to_fleet(fleet_slug)


# Initialize the application
def init_app():
    """Initialize the application with default files"""
    with app.app_context():
        db.create_all()
        print(f"✅ Database ready at {app.config['SQLALCHEMY_DATABASE_URI']}")
        setup_default_files()
        print(f"🔹 DEBUG: Current Admin Password: {PASSWORD}")
        print(f"✅ Application initialized successfully")


@app.before_first_request
def ensure_setup():
    """Ensure the application is fully initialised before serving requests."""
    init_app()


# Run the application
if __name__ == "__main__":
    init_app()
    app.run(debug=os.getenv("FLASK_DEBUG", "False") == "True")
