import os
from datetime import datetime
import pandas as pd
import requests
from flask import (
    Flask, render_template, request, redirect,
    url_for, session, flash, send_file
)
from werkzeug.utils import secure_filename
from dotenv import load_dotenv
from functools import wraps
from flask_sqlalchemy import SQLAlchemy
from flask_migrate import Migrate
from flask_caching import Cache
from sqlalchemy import create_engine


# Load environment variables
load_dotenv()

# Base directory
BASE_DIR = os.getcwd()
DATA_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)

# Flask App
app = Flask(__name__, template_folder="Templates")
app.secret_key = os.getenv("FLASK_SECRET_KEY", "fallback_secret_key")

# Cache
cache = Cache(app, config={"CACHE_TYPE": "simple"})

# Password
PASSWORD = os.getenv("ADMIN_PASSWORD", "Admin.123.")

# File Paths
EMPLOYEE_FILE = os.path.join(DATA_DIR, "employees.csv")
COMPANY_FILE = os.path.join(DATA_DIR, "companies.csv")
PROJECT_FILE = os.path.join(DATA_DIR, "projects.csv")
HARM_FILE = os.path.join(DATA_DIR, "harm_drive_vehicles.csv")
WAHOO_FILE = os.path.join(DATA_DIR, "wahoo_pool_vehicles.csv")

# Allowed extensions
ALLOWED_EXTENSIONS = {"pdf", "docx", "jpg", "jpeg", "png"}

# Util
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

def calculate_expiry(rego_date):
    try:
        today = datetime.today()
        rego_date = datetime.strptime(rego_date, '%Y-%m-%d')
        return (rego_date - today).days
    except Exception:
        return None

@app.template_filter('datetimeformat')
def datetimeformat(value):
    try:
        return datetime.strptime(value, '%Y-%m-%d').strftime('%d/%m/%Y')
    except Exception:
        return value

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("logged_in"):
            flash("Please log in first.", "warning")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        if request.form["password"] == PASSWORD:
            session["logged_in"] = True
            flash("Login successful", "success")
            return redirect(url_for("index"))
        else:
            flash("Wrong password", "danger")
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    flash("Logged out", "info")
    return redirect(url_for("login"))

@app.route("/")
@login_required
def index():
    try:
        vehicles_df = pd.concat([
            pd.read_csv(WAHOO_FILE),
            pd.read_csv(HARM_FILE)
        ], ignore_index=True)
        vehicles = vehicles_df.to_dict(orient="records")
        for v in vehicles:
            v["Expiry"] = calculate_expiry(v.get("Rego Renewal Date", ""))
    except Exception as e:
        flash(f"Error reading vehicle data: {e}", "danger")
        vehicles = []

    return render_template("index.html", vehicles=vehicles)

@app.route("/employees")
@login_required
def employees():
    try:
        df = pd.read_csv(EMPLOYEE_FILE)
        employees = df.to_dict(orient="records")
        return render_template("employees.html", employees=employees)
    except Exception as e:
        flash(f"Error loading employees: {e}", "danger")
        return redirect(url_for("index"))

@app.route("/companies")
@login_required
def companies():
    try:
        df = pd.read_csv(COMPANY_FILE)
        companies = df.to_dict(orient="records")
        return render_template("companies.html", companies=companies)
    except Exception as e:
        flash(f"Error loading companies: {e}", "danger")
        return redirect(url_for("index"))

@app.route("/projects")
@login_required
def projects():
    try:
        df = pd.read_csv(PROJECT_FILE)
        projects = df.to_dict(orient="records")
        return render_template("projects.html", projects=projects)
    except Exception as e:
        flash(f"Error loading projects: {e}", "danger")
        return redirect(url_for("index"))


# Logo path setup
logo_filename = "default_logo.png"
logo_path = os.path.join(BASE_DIR, "static", "logos", logo_filename)
if os.path.exists(logo_path):
    print(f"✅ Logo found: {logo_path}")

print(f"🔹 DEBUG: Current Admin Password: {PASSWORD}")

# Database Models
class Company(db.Model):
    __tablename__ = 'companies'
    
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(255), nullable=False)
    registration_date = db.Column(db.DateTime)
    abn = db.Column(db.String(20), unique=True)
    acn = db.Column(db.String(20))
    company_type = db.Column(db.String(50))
    registered_address = db.Column(db.Text)
    qbcc_license_number = db.Column(db.String(50))
    document_folder = db.Column(db.String(255))
    
    def __repr__(self):
        return f"<Company {self.name}>"

class Employee(db.Model):
    __tablename__ = 'employees'
    
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    family_name = db.Column(db.String(100), nullable=False)
    tfn = db.Column(db.String(20))
    abn = db.Column(db.String(20))
    address = db.Column(db.Text)
    email = db.Column(db.String(120))
    phone = db.Column(db.String(20))
    
    def __repr__(self):
        return f"<Employee {self.name} {self.family_name}>"

class Project(db.Model):
    __tablename__ = 'projects'
    
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(255), nullable=False, unique=True)
    description = db.Column(db.Text)
    start_date = db.Column(db.DateTime)
    end_date = db.Column(db.DateTime)
    image_filename = db.Column(db.String(255))
    
    def __repr__(self):
        return f"<Project {self.name}>"

class Vehicle(db.Model):
    __tablename__ = 'vehicles'
    
    id = db.Column(db.Integer, primary_key=True)
    plate = db.Column(db.String(20), nullable=False, unique=True)
    type = db.Column(db.String(50))
    vin = db.Column(db.String(50))
    rego_renewal_date = db.Column(db.String(10))
    insurance_renewal_date = db.Column(db.String(10))
    expiry = db.Column(db.Integer)
    value = db.Column(db.Float)
    transfer_fee = db.Column(db.Float)
    
    def __repr__(self):
        return f"<Vehicle {self.plate}>"
    
    @property
    def calculated_expiry(self):
        """Dynamically calculate days until registration expiry"""
        return calculate_expiry(self.rego_renewal_date)

# Utility functions
def allowed_file(filename):
    """
    Check if a file has an allowed extension
    """
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

def calculate_expiry(rego_date):
    """
    Calculate the number of days until a given date.
    :param rego_date: Registration renewal date as a string in 'YYYY-MM-DD' format.
    :return: Number of days until the date or None if invalid.
    """
    try:
        today = datetime.today()
        rego_date = datetime.strptime(rego_date, '%Y-%m-%d')
        return (rego_date - today).days
    except Exception:
        return None

def format_date_ddmmyyyy(date):
    """
    Format a date object to DD/MM/YYYY string format
    """
    if pd.notnull(date):
        return datetime.strptime(str(date), "%Y-%m-%d").strftime("%d/%m/%Y")
    return None

# Middleware to protect routes
def login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("logged_in"):
            flash("You must log in to access this page.", "error")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return wrapper

# Template filter to format date strings
@app.template_filter('datetimeformat')
def datetimeformat(value):
    try:
        return datetime.strptime(value, '%Y-%m-%d').strftime('%d/%m/%Y')
    except Exception:
        return value

# Setup functions
def setup_directories():
    """Ensure all required directories exist"""
    os.makedirs(PROJECT_IMAGES_DIR, exist_ok=True)
    os.makedirs(MAIN_DIR, exist_ok=True)
    os.makedirs(os.path.join(BASE_DIR, "static", "logos"), exist_ok=True)

def setup_default_files():
    """Create default Excel files if they don't exist"""
    # Create default projects file if not exists
    if not os.path.exists(PROJECTS_FILE):
        pd.DataFrame(columns=["Project Name", "Description", "Start Date", "End Date", "Image Filename"]).to_excel(
            PROJECTS_FILE, index=False, engine="openpyxl")
        print(f"✅ Created default projects file: {PROJECTS_FILE}")
        
    # Create default employees file if not exists
    if not os.path.exists(EMPLOYEE_FILE):
        pd.DataFrame(columns=["Name", "Family Name", "TFN", "ABN", "Address", "Email", "Phone"]).to_excel(
            EMPLOYEE_FILE, index=False, engine="openpyxl")
        print(f"✅ Created default employees file: {EMPLOYEE_FILE}")
        
    # Create default companies file if not exists
    if not os.path.exists(EXCEL_FILE):
        pd.DataFrame(columns=["Company Name", "Registration Date", "ABN", "ACN", "Type", "Registered Address", 
                             "QBCC License Number", "Documents"]).to_excel(
            EXCEL_FILE, index=False, engine="openpyxl")
        print(f"✅ Created default companies file: {EXCEL_FILE}")
        
    # Setup Harm Drive data file
    setup_harm_drive_file()
    
    # Setup Wahoo vehicles file
    setup_wahoo_vehicles_file()

def setup_harm_drive_file():
    """Create HarmDriveData.xlsx if it doesn't exist"""
    if not os.path.exists(HARM_DRIVE_FILE):
        data = {
            "Plate": ["371RLI", "295RMK", "975SMU"],
            "Type": ["Triton", "BT-50", "Triton"],
            "VIN": ["MMAENKA40BD006265", "MM0UNY0W400891903", "MMAJNKB40CD018749"],
            "Rego Renewal Date": ["2025-03-25", "2025-03-06", "2025-03-20"],
            "Insurance Renewal (CTP) Date": ["", "", ""],
            "Expiry": [105, 86, 100],
            "Value": [10000, 12500, 12500],
            "Transfer Fee": [None, 406.5, 406.5]
        }
        df = pd.DataFrame(data)
        df.to_excel(HARM_DRIVE_FILE, index=False, engine="openpyxl")
        print(f"✅ Created HarmDriveData.xlsx file")

def setup_wahoo_vehicles_file():
    """Create wahoo_pool_vehicles.xlsx if it doesn't exist"""
    if not os.path.exists(WAHOO_VEHICLES_FILE):
        required_columns = ["Plate", "Type", "VIN", "Rego Renewal Date", "Insurance Renewal (CTP) Date", 
                           "Value", "Transfer Fee", "Expiry"]
        df = pd.DataFrame(columns=required_columns)
        df.to_excel(WAHOO_VEHICLES_FILE, index=False, engine="openpyxl")
        print(f"✅ Created Wahoo pool vehicles file")


@cache.cached(timeout=600)  # Cache for 10 minutes
def get_qld_construction_news():
    """
    Fetch recent Queensland construction news from News API
    """
    # Load API key securely (replace with a default fallback if needed)
    api_key = os.getenv('NEWS_API_KEY', 'bfb864ec86be43f49b257cb04ff2ab0f')

    # Endpoint for the News API
    endpoint = 'https://newsapi.org/v2/everything'

    # Define search parameters
    params = {
        'q': 'Queensland construction',
        'language': 'en',
        'sortBy': 'publishedAt',
        'pageSize': 20,  # Limit results to the first 20 articles for efficiency
        'apiKey': api_key
    }

    try:
        # Make the API request
        response = requests.get(endpoint, params=params)
        response.raise_for_status()  # Raise HTTPError for bad status codes

        # Parse JSON response
        data = response.json()

        # Extract relevant fields for articles
        articles = data.get('articles', [])
        filtered_articles = []
        for article in articles:
            filtered_articles.append({
                'title': article.get('title', 'No title'),
                'url': article.get('url', '#'),
                'publishedAt': article.get('publishedAt', 'Unknown date')
            })

        return filtered_articles  # Return cleaned list of articles

    except requests.exceptions.RequestException as e:
        # Handle connection or request-related errors
        print(f"Error fetching news: {e}")
        return []
    except ValueError as e:
        # Handle JSON parsing errors
        print(f"Error parsing response JSON: {e}")
        return []

# Template utility for company logos
@app.template_global()
def get_company_logo_static(company_name, abn):
    """Returns the static URL for a company logo or a default image."""
    logo_filename = f"{company_name.replace(' ', '_')}_{abn}.png"  # Replace spaces with underscores
    logo_path = os.path.join("static", "logos", logo_filename)

    print(f"🔍 Checking for logo: {logo_path}")  # Debugging output

    if os.path.exists(os.path.join(BASE_DIR, logo_path)):
        return url_for("static", filename=f"logos/{logo_filename}")
    else:
        print(f"⚠️ Logo not found, using default logo.")  # Debugging output
        return url_for("static", filename="logos/default_logo.png")

# Routes
@app.route("/navbar")
def navbar():
    """Serve the navbar template"""
    return render_template("navbar.html")

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
    return render_template("login.html", title="Login")

@app.route("/logout")
def logout():
    """Handle user logout"""
    session.clear()
    flash("You have been logged out.", "success")
    return redirect(url_for("login"))

@app.route("/")
@login_required
def index():
    """Main dashboard page"""
    try:
        # Load vehicles with registration dates nearing expiry
        vehicles = Vehicle.query.all()
        
        # Get construction news
        news_items = get_qld_construction_news()

        # Pagination Logic
        PER_PAGE = 5  # 5 news items per page
        page = request.args.get('page', 1, type=int)  # Get page from query params
        total_pages = (len(news_items) + PER_PAGE - 1) // PER_PAGE
        page = max(1, min(page, total_pages))  # Ensure page is within valid range
        paginated_news = news_items[(page - 1) * PER_PAGE: page * PER_PAGE]

        return render_template("index.html", vehicles=vehicles, news_items=paginated_news, page=page, total_pages=total_pages)
    except Exception as e:
        flash(f"Error loading data: {e}", "error")
        return render_template("index.html", vehicles=[], news_items=[], page=1, total_pages=1)

@app.route("/employees")
@login_required
def employees():
    """Display list of employees"""
    try:
        all_employees = Employee.query.all()
        return render_template("employees.html", employees=all_employees)
    except Exception as e:
        flash(f"Error loading employees: {e}", "error")
        return redirect(url_for("index"))

@app.route("/search_employee", methods=["GET", "POST"])
@login_required
def search_employee():
    """Search for an employee by name and family name"""
    selected_employees = []

    if request.method == "POST":
        search_name = request.form.get("search_name", "").strip().lower()
        search_family_name = request.form.get("search_family_name", "").strip().lower()

        # Ensure both name and family name are provided
        if not search_name or not search_family_name:
            flash("Both Name and Family Name are required!", "error")
            return render_template("search_employee.html", selected_employees=[])

        # Find matching employees in database
        matching_employees = Employee.query.filter(
            db.func.lower(Employee.name) == search_name,
            db.func.lower(Employee.family_name) == search_family_name
        ).all()
        
        if matching_employees:
            selected_employees = matching_employees
        else:
            flash("No employees found with the specified criteria.", "warning")

    return render_template("search_employee.html", selected_employees=selected_employees)

@app.route("/add_employee", methods=["GET", "POST"])
@login_required
def add_employee():
    """Add a new employee"""
    if request.method == "POST":
        try:
            # Create new employee record
            new_employee = Employee(
                name=request.form["name"].strip(),
                family_name=request.form["family_name"].strip(),
                tfn=request.form["tfn"].strip(),
                abn=request.form["abn"].strip(),
                address=request.form["address"].strip(),
                email=request.form["email"].strip(),
                phone=request.form["phone"].strip()
            )
            
            # Add to database
            db.session.add(new_employee)
            db.session.commit()
            
            flash("Employee added successfully!", "success")
            return redirect(url_for("employees"))
        except Exception as e:
            db.session.rollback()
            flash(f"Error adding employee: {e}", "error")
            return redirect(url_for("add_employee"))

    return render_template("add_employee.html")

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

            # Check if company already exists
            existing_company = Company.query.filter_by(abn=abn).first()
            if existing_company:
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

            # ✅ Save the ASIC Extract file
            asic_extract.save(os.path.join(folder_path, secure_filename(asic_extract.filename)))

            # Handle optional document uploads
            for field_name in ["company_registration", "logo"]:
                file = request.files.get(field_name)
                if file and file.filename and allowed_file(file.filename):
                    file.save(os.path.join(folder_path, secure_filename(file.filename)))


            # Check if company already exists
            existing_company = Company.query.filter_by(abn=abn).first()
            if existing_company:
                flash("A company with this ABN already exists!", "error")
                return redirect(url_for("add_company"))

            # Create new company record in database
            new_company = Company(
                name=company_name,
                registration_date=datetime.strptime(registration_date, '%Y-%m-%d'),
                abn=abn,
                acn=acn,
                company_type=company_type,
                registered_address=registered_address,
                qbcc_license_number=qbcc_license_number,
                document_folder=folder_name
            )
            db.session.add(new_company)
            db.session.commit()

            flash("Company added successfully!", "success")
            return redirect(url_for("view_company"))

        # Render the Add Company form
        return render_template("add_company.html", title="Add Company")

    except Exception as e:
        app.logger.error(f"Error adding company: {e}")
        flash(f"Error adding company: {e}", "error")
        return redirect(url_for("view_company"))



@app.route("/view_company", methods=["GET", "POST"])
@login_required
def view_company():
    try:
        df = pd.read_excel(EXCEL_FILE, engine="openpyxl", dtype={"QBCC License Number": str})
        if df.empty:
            flash("No companies found. Please add a company first.", "info")
            return redirect(url_for("add_company"))

        companies = df.to_dict(orient="records")

        if request.method == "POST":
            selected_company = request.form.get("company_name")
            if not selected_company:
                flash("Please select a company.", "error")
                return redirect(url_for("view_company"))

            company = df[df["Company Name"].str.strip() == selected_company.strip()]
            if company.empty:
                flash("Company not found.", "error")
                return redirect(url_for("view_company"))

            company = company.iloc[0].to_dict()

            folder_name = company["Documents"]
            folder_path = os.path.join(MAIN_DIR, folder_name)
            if not os.path.exists(folder_path):
                flash("Document folder does not exist.", "error")
                return redirect(url_for("view_company"))

            # Debugging Output
            print(f"📌 Selected company from form: {selected_company}")
            print(f"🔍 Matching companies found: {company}")

            logo_url = get_company_logo_static(company["Company Name"], company["ABN"])

            documents = {doc: doc for doc in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, doc))}

            return render_template(
                "company_details.html",
                title="Company Details",
                company=company,
                documents=documents,
                logo_url=logo_url
            )

        return render_template("view_company.html", title="View Companies", companies=companies)

    except Exception as e:
        flash(f"Error loading companies: {e}", "error")
        print(f"[ERROR] {e}")
        return redirect(url_for("index"))



@app.route("/add_project", methods=["GET", "POST"])
@login_required
def add_project():
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

            # Handle image upload
            image = request.files.get("project_image")
            if image and allowed_file(image.filename):
                filename = secure_filename(image.filename)
                image.save(os.path.join(PROJECT_IMAGES_DIR, filename))
            else:
                filename = None  # No image uploaded or invalid file

            # Prepare project data for saving
            new_project = {
                "Project Name": project_name,
                "Description": project_description,
                "Start Date": start_date,
                "End Date": end_date,
                "Image Filename": filename
            }

            # Append project data to the Excel file
            if not os.path.exists(PROJECTS_FILE):
                # Create file if it does not exist
                pd.DataFrame([new_project]).to_excel(PROJECTS_FILE, index=False, engine="openpyxl")
            else:
                df = pd.read_excel(PROJECTS_FILE, engine="openpyxl")
                df = pd.concat([df, pd.DataFrame([new_project])], ignore_index=True)
                df.to_excel(PROJECTS_FILE, index=False, engine="openpyxl")

            flash("Project added successfully!", "success")
            return redirect(url_for("view_projects"))
        except Exception as e:
            flash(f"Error adding project: {str(e)}", "error")
            return redirect(url_for("add_project"))

    return render_template("add_project.html")

@app.route("/view_projects")
@login_required
def view_projects():
    try:
        if not os.path.exists(PROJECTS_FILE):
            flash("No projects found. Please add a project first.", "info")
            return redirect(url_for("add_project"))

        df = pd.read_excel(PROJECTS_FILE, engine="openpyxl")
        projects = df.to_dict(orient="records")
        return render_template("view_projects.html", projects=projects)
    except Exception as e:
        flash(f"Error loading projects: {str(e)}", "error")
        return redirect(url_for("index"))


@app.route("/view_project/<project_name>")
@login_required
def view_project(project_name):
    try:
        if not os.path.exists(PROJECTS_FILE):
            flash("Projects data file does not exist!", "error")
            return redirect(url_for("view_projects"))

        df = pd.read_excel(PROJECTS_FILE, engine="openpyxl")

        # Retrieve the project matching the given project name
        project_data = df[df["Project Name"] == project_name].to_dict(orient="records")
        if not project_data:
            flash(f"Project '{project_name}' not found.", "error")
            return redirect(url_for("view_projects"))

        # Pass the first matched project (should only be one)
        project = project_data[0]
        return render_template("view_project.html", project=project)
    except Exception as e:
        flash(f"Error retrieving project: {str(e)}", "error")
        return redirect(url_for("view_projects"))



@app.template_global()
def get_company_logo_static(company_name, abn):
    """Returns the static URL for a company logo or a default image."""
    logo_filename = f"{company_name.replace(' ', '_')}_{abn}.png"  # Replace spaces with underscores
    logo_path = os.path.join("static", "logos", logo_filename)

    print(f"🔍 Checking for logo: {logo_path}")  # Debugging output

    if os.path.exists(logo_path):
        return url_for("static", filename=f"logos/{logo_filename}")
    else:
        print(f"⚠️ Logo not found, using default logo.")  # Debugging output
        return url_for("static", filename="default_logo.png")


@app.route("/download/<company_folder>/<filename>")
def download_file(company_folder, filename):
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



if __name__ == "__main__":
    app.run(debug=os.getenv("FLASK_DEBUG", "False") == "True")

