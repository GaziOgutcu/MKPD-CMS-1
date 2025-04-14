import os
from datetime import datetime
import pandas as pd
from flask import (
    Flask, render_template, request, redirect,
    url_for, session, flash, send_file
)
from werkzeug.utils import secure_filename
from dotenv import load_dotenv
from functools import wraps
from flask_caching import Cache
import requests

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
EXCEL_FILE = os.path.join(DATA_DIR, "companies.xlsx")  # Added missing EXCEL_FILE
PROJECTS_FILE = os.path.join(DATA_DIR, "projects.xlsx")  # Added missing PROJECTS_FILE
HARM_DRIVE_FILE = os.path.join(DATA_DIR, "HarmDriveData.xlsx")  # Added missing HARM_DRIVE_FILE
WAHOO_VEHICLES_FILE = os.path.join(DATA_DIR, "wahoo_pool_vehicles.xlsx")  # Added missing WAHOO_VEHICLES_FILE

# Flask App configuration
app = Flask(__name__, template_folder="Templates")
app.secret_key = os.getenv("FLASK_SECRET_KEY", "fallback_secret_key")
cache = Cache(app, config={"CACHE_TYPE": "simple"})

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
        rego_date = datetime.strptime(rego_date, '%Y-%m-%d')
        return (rego_date - today).days
    except Exception:
        return None

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
    return render_template("login.html", title="Login")

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
        
        # Get construction news
        news_items = get_qld_construction_news()

        # Pagination for news
        PER_PAGE = 5
        page = request.args.get('page', 1, type=int)
        total_pages = (len(news_items) + PER_PAGE - 1) // PER_PAGE
        page = max(1, min(page, total_pages))
        paginated_news = news_items[(page - 1) * PER_PAGE: page * PER_PAGE]

        return render_template("index.html", vehicles=vehicles, news_items=paginated_news, 
                              page=page, total_pages=total_pages)
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
        df = pd.read_csv(EMPLOYEE_FILE)
        employees_list = df.to_dict(orient="records")
        return render_template("employees.html", employees=employees_list)
    except Exception as e:
        flash(f"Error loading employees: {e}", "danger")
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

        try:
            # Find matching employees
            df = pd.read_csv(EMPLOYEE_FILE)
            matching_employees = df[
                (df["Name"].str.lower() == search_name) & 
                (df["Family Name"].str.lower() == search_family_name)
            ]
            
            if not matching_employees.empty:
                selected_employees = matching_employees.to_dict(orient="records")
            else:
                flash("No employees found with the specified criteria.", "warning")
        except Exception as e:
            flash(f"Error searching employees: {e}", "error")

    return render_template("search_employee.html", selected_employees=selected_employees)

@app.route("/add_employee", methods=["GET", "POST"])
@login_required
def add_employee():
    """Add a new employee"""
    if request.method == "POST":
        try:
            new_emp = {
                "Name": request.form["name"].strip(),
                "Family Name": request.form["family_name"].strip(),
                "TFN": request.form["tfn"].strip(),
                "ABN": request.form["abn"].strip(),
                "Address": request.form["address"].strip(),
                "Email": request.form["email"].strip(),
                "Phone": request.form["phone"].strip()
            }
            
            if os.path.exists(EMPLOYEE_FILE):
                df = pd.read_csv(EMPLOYEE_FILE)
                df = pd.concat([df, pd.DataFrame([new_emp])], ignore_index=True)
            else:
                df = pd.DataFrame([new_emp])
                
            df.to_csv(EMPLOYEE_FILE, index=False)
            flash("Employee added successfully!", "success")
            return redirect(url_for("employees"))
        except Exception as e:
            flash(f"Error saving employee: {e}", "danger")
            return redirect(url_for("add_employee"))
            
    return render_template("add_employee.html")

@app.route("/companies")
@login_required
def companies():
    """Display list of companies"""
    try:
        df = pd.read_csv(COMPANY_FILE)
        companies_list = df.to_dict(orient="records")
        return render_template("companies.html", companies=companies_list)
    except Exception as e:
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

            # Check if company exists (Excel file)
            try:
                df = pd.read_excel(EXCEL_FILE, engine="openpyxl")
                if not df[df["ABN"] == abn].empty:
                    flash("A company with this ABN already exists!", "error")
                    return redirect(url_for("add_company"))
            except Exception:
                # File might not exist yet
                df = pd.DataFrame(columns=["Company Name", "Registration Date", "ABN", "ACN", 
                                         "Type", "Registered Address", "QBCC License Number", "Documents"])

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

            # Add to Excel file
            new_company = {
                "Company Name": company_name,
                "Registration Date": registration_date,
                "ABN": abn,
                "ACN": acn,
                "Type": company_type,
                "Registered Address": registered_address,
                "QBCC License Number": qbcc_license_number,
                "Documents": folder_name
            }
            
            df = pd.concat([df, pd.DataFrame([new_company])], ignore_index=True)
            df.to_excel(EXCEL_FILE, index=False, engine="openpyxl")

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
    """View company details"""
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

@app.route("/projects")
@login_required
def projects():
    """Display list of projects"""
    try:
        df = pd.read_csv(PROJECT_FILE)
        projects_list = df.to_dict(orient="records")
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

            # Handle image upload
            image = request.files.get("project_image")
            if image and allowed_file(image.filename):
                filename = secure_filename(image.filename)
                image.save(os.path.join(PROJECT_IMAGE_DIR, filename))
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

            # Save to both CSV and Excel files for compatibility
            # Excel file
            if not os.path.exists(PROJECTS_FILE):
                pd.DataFrame([new_project]).to_excel(PROJECTS_FILE, index=False, engine="openpyxl")
            else:
                df = pd.read_excel(PROJECTS_FILE, engine="openpyxl")
                df = pd.concat([df, pd.DataFrame([new_project])], ignore_index=True)
                df.to_excel(PROJECTS_FILE, index=False, engine="openpyxl")
                
            # CSV file
            if not os.path.exists(PROJECT_FILE):
                pd.DataFrame([new_project]).to_csv(PROJECT_FILE, index=False)
            else:
                df = pd.read_csv(PROJECT_FILE)
                df = pd.concat([df, pd.DataFrame([new_project])], ignore_index=True)
                df.to_csv(PROJECT_FILE, index=False)

            flash("Project added successfully!", "success")
            return redirect(url_for("view_projects"))
        except Exception as e:
            flash(f"Error adding project: {str(e)}", "error")
            return redirect(url_for("add_project"))

    return render_template("add_project.html")

@app.route("/view_projects")
@login_required
def view_projects():
    """View all projects"""
    try:
        if os.path.exists(PROJECTS_FILE):
            df = pd.read_excel(PROJECTS_FILE, engine="openpyxl")
        elif os.path.exists(PROJECT_FILE):
            df = pd.read_csv(PROJECT_FILE)
        else:
            flash("No projects found. Please add a project first.", "info")
            return redirect(url_for("add_project"))
            
        projects = df.to_dict(orient="records")
        return render_template("view_projects.html", projects=projects)
    except Exception as e:
        flash(f"Error loading projects: {str(e)}", "error")
        return redirect(url_for("index"))

@app.route("/view_project/<project_name>")
@login_required
def view_project(project_name):
    """View details of a specific project"""
    try:
        if os.path.exists(PROJECTS_FILE):
            df = pd.read_excel(PROJECTS_FILE, engine="openpyxl")
        elif os.path.exists(PROJECT_FILE):
            df = pd.read_csv(PROJECT_FILE)
        else:
            flash("Projects data file does not exist!", "error")
            return redirect(url_for("view_projects"))

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

# Initialize the application
def init_app():
    """Initialize the application with default files"""
    setup_default_files()
    print(f"🔹 DEBUG: Current Admin Password: {PASSWORD}")
    print(f"✅ Application initialized successfully")

# Run the application
if __name__ == "__main__":
    init_app()
    app.run(debug=os.getenv("FLASK_DEBUG", "False") == "True")