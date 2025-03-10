import os
from datetime import datetime
import pandas as pd
import requests
from flask import Flask, render_template, send_from_directory, request, redirect, url_for, session, flash, send_file, jsonify
from werkzeug.utils import secure_filename
from dotenv import load_dotenv
from functools import wraps
from flask_sqlalchemy import SQLAlchemy
from flask_migrate import Migrate
from flask_caching import Cache

# Define Base Directory
BASE_DIR = os.getcwd()

# Initialize Flask app
app = Flask(__name__, template_folder='Templates')
app.secret_key = os.getenv("FLASK_SECRET_KEY", "fallback_secret_key")

# print(f"Secret Key: {app.secret_key}")

# print(app.jinja_loader.searchpath)

PROJECTS_FILE = os.path.join(BASE_DIR, "projects.xlsx")

REQUIRED_ENV_VARS = ["FLASK_SECRET_KEY", "ADMIN_PASSWORD", "DATABASE_URL"]
for var in REQUIRED_ENV_VARS:
    if not os.getenv(var):
        raise ValueError(f"❌ Missing environment variable: {var}")


# Secure Environment Variables
PASSWORD = os.getenv("ADMIN_PASSWORD", "default_fallback_password")


# Allowed File Extensions
ALLOWED_EXTENSIONS = {"pdf", "docx", "jpg", "jpeg", "png"}

# Project Images Directory
PROJECT_IMAGES_DIR = os.path.join(BASE_DIR, "static", "project_images")
os.makedirs(PROJECT_IMAGES_DIR, exist_ok=True)  # Ensure it exists

# File path for Wahoo Pool Vehicles
WAHOO_VEHICLES_FILE = "wahoo_pool_vehicles.xlsx"

# Load environment variables
load_dotenv()


# ✅ Configure PostgreSQL Database for Railway
DATABASE_URL = os.getenv("DATABASE_URL")
if not DATABASE_URL:
    raise ValueError("❌ DATABASE_URL is not set! Make sure it's configured in Railway.")

# ✅ Fix PostgreSQL SSL Mode Issue
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://")

app.config["SQLALCHEMY_DATABASE_URI"] = DATABASE_URL
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

#NEWS FETCH CACHE
cache = Cache(app, config={"CACHE_TYPE": "simple"})

# ✅ Initialize Database
db = SQLAlchemy(app)
migrate = Migrate(app, db)

# Secure Environment Variables
PASSWORD = os.getenv("ADMIN_PASSWORD", "default_fallback_password")

# ✅ Database Model (Replaces Excel)
class Vehicle(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    plate = db.Column(db.String(10), unique=True, nullable=False)
    type = db.Column(db.String(50), nullable=False)
    vin = db.Column(db.String(50), unique=True, nullable=False)
    rego_renewal_date = db.Column(db.String(20), nullable=False)
    insurance_renewal_date = db.Column(db.String(20), nullable=True)
    expiry = db.Column(db.Integer, nullable=True)
    value = db.Column(db.Float, nullable=True)
    transfer_fee = db.Column(db.Float, nullable=True)

    def __repr__(self):
        return f"<Vehicle {self.plate}>"

# Define Base Directory
MAIN_DIR = os.path.join(BASE_DIR, "CompanyFolders")
os.makedirs(MAIN_DIR, exist_ok=True)  # Ensure it exists


logo_filename = "default_logo.png"  # Ensure it has a default value
logo_path = os.path.join(BASE_DIR, "static", "logos", logo_filename)

if os.path.exists(logo_path):
    print(f"✅ Logo found: {logo_path}")  # ✅ Fixed



if not os.path.exists(PROJECTS_FILE):
    pd.DataFrame(columns=["Project Name", "Description", "Start Date", "End Date"]).to_excel(PROJECTS_FILE, index=False)


print(f"🔹 DEBUG: Current Admin Password: {PASSWORD}")


# Utility to check allowed file extensions
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

# Middleware to protect routes
def login_required(f):
    @wraps(f)  # 🔹 Preserves function metadata
    def wrapper(*args, **kwargs):
        if not session.get("logged_in"):
            flash("You must log in to access this page.", "error")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return wrapper

def calculate_expiry(rego_date):
    """
    Calculate the number of days until a given date.
    :param rego_date: Rego renewal date as a string in 'YYYY-MM-DD' format.
    :return: Number of days until the date or None if invalid.
    """
    try:
        today = datetime.today()
        rego_date = datetime.strptime(rego_date, '%Y-%m-%d')
        return (rego_date - today).days
    except Exception:
        return None

def format_date_ddmmyyyy(date):
    if pd.notnull(date):
        return datetime.strptime(str(date), "%Y-%m-%d").strftime("%d/%m/%Y")
    return None


@cache.cached(timeout=600)  # Cache for 10 minutes
def get_qld_construction_news():
    # Load API key securely (replace with a default fallback if needed)
    api_key = os.getenv('NEWS_API_KEY', 'bfb864ec86be43f49b257cb04ff2ab0f')  # Secure environment variable

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

# ✅ Route to serve the Navbar file
@app.route("/navbar")
def navbar():
    return render_template("navbar.html")


@app.template_filter('datetimeformat')
def datetimeformat(value):
    """
    Convert a date string from 'YYYY-MM-DD' to 'DD/MM/YYYY'.
    """
    try:
        return datetime.strptime(value, '%Y-%m-%d').strftime('%d/%m/%Y')
    except Exception:
        return value  # If the date is invalid, return the original value



@app.route("/login", methods=["GET", "POST"])
def login():
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
    session.clear()
    flash("You have been logged out.", "success")
    return redirect(url_for("login"))

@app.route("/")
@login_required
def index():
    try:
        # Initialize variables
        vehicles = []
        news_items = get_qld_construction_news()

        # Vehicles logic (unchanged)
        # ✅ Fetch all vehicles from PostgreSQL
        vehicles = Vehicle.query.all()


        # Pagination Logic
        PER_PAGE = 5  # 5 news items per page
        page = page = max(1, min(page, total_pages))  # Ensure page is within valid range
        total_pages = (len(news_items) + PER_PAGE - 1) // PER_PAGE
        paginated_news = news_items[(page - 1) * PER_PAGE: page * PER_PAGE]

        return render_template("index.html", vehicles=vehicles, news_items=paginated_news, page=page, total_pages=total_pages)

    except Exception as e:
        flash(f"Error loading data: {e}", "error")
        return render_template("index.html", vehicles=[], news_items=[], page=1, total_pages=1)



@app.route("/CompanyFolders/<path:filename>")
def serve_company_file(filename):
    """Serve files from the CompanyFolders directory."""
    return send_from_directory(MAIN_DIR, filename)
def setup_harm_drive_file():
    harm_drive_file = os.path.join(os.getcwd(), "HarmDriveData.xlsx")
    if not os.path.exists(harm_drive_file):
        data = {
            "Plate": ["371RLI", "295RMK", "975SMU"],
            "Type": ["Triton", "BT-50", "Triton"],
            "VIN": ["MMAENKA40BD006265", "MM0UNY0W400891903", "MMAJNKB40CD018749"],
            "Rego Renewal Date": ["25/03/2025", "06/03/2025", "20/03/2025"],
            "Insurance Renewal (CTP) Date": ["", "", ""],
            "Expiry": [105, 86, 100],
            "Value": [10000, 12500, 12500],
            "Transfer Fee": [None, 406.5, 406.5]
        }
        df = pd.DataFrame(data)
        df.to_excel(harm_drive_file, index=False, engine="openpyxl")
        print("HarmDriveData.xlsx created.")


# ✅ Fetch All Vehicles (Replaces Pandas Read Excel)
@app.route("/harm_drive")
@login_required
def harm_drive():
    vehicles = Vehicle.query.all()  # Fetch all vehicles from PostgreSQL
    return render_template("harm_drive.html", vehicles=vehicles)

# ✅ Add New Vehicle (Replaces Pandas Excel Handling)
@app.route("/add_vehicle", methods=["POST"])
@login_required
def add_vehicle():
    try:
        new_vehicle = Vehicle(
            plate=request.form["plate"].strip(),
            type=request.form["type"].strip(),
            vin=request.form["vin"].strip(),
            rego_renewal_date=request.form["rego_renewal_date"].strip(),
            insurance_renewal_date=request.form.get("ctp_date", "").strip(),
            value=float(request.form["value"]),
            transfer_fee=float(request.form.get("transfer_fee", 0))
        )
        db.session.add(new_vehicle)
        db.session.commit()
        flash("Vehicle added successfully!", "success")
    except Exception as e:
        flash(f"Error adding vehicle: {e}", "error")

    return redirect(url_for("harm_drive"))

# ✅ Update a Vehicle’s Registration Date
@app.route("/update_vehicle", methods=["POST"])
@login_required
def update_vehicle():
    try:
        plate_to_update = request.form["plate_to_update"].strip()
        new_date = request.form["new_rego_renewal_date"].strip()

        vehicle = Vehicle.query.filter_by(plate=plate_to_update).first()
        if vehicle:
            vehicle.rego_renewal_date = new_date
            db.session.commit()
            flash(f"Vehicle {plate_to_update} updated successfully!", "success")
        else:
            flash(f"Vehicle {plate_to_update} not found!", "error")
    except Exception as e:
        flash(f"Error updating vehicle: {e}", "error")

    return redirect(url_for("harm_drive"))

# ✅ Delete a Vehicle
@app.route("/delete_vehicle", methods=["POST"])
@login_required
def delete_vehicle():
    try:
        plate_to_delete = request.form["plate_to_delete"].strip()
        vehicle = Vehicle.query.filter_by(plate=plate_to_delete).first()

        if vehicle:
            db.session.delete(vehicle)
            db.session.commit()
            flash(f"Vehicle {plate_to_delete} deleted successfully!", "success")
        else:
            flash(f"Vehicle {plate_to_delete} not found!", "error")
    except Exception as e:
        flash(f"Error deleting vehicle: {e}", "error")

    return redirect(url_for("harm_drive"))


# Ensure the Wahoo Vehicles file exists
def setup_wahoo_vehicles_file():
    if not os.path.exists(WAHOO_VEHICLES_FILE):
        required_columns = ["Plate", "Type", "VIN", "Rego Renewal Date", "Insurance Renewal (CTP) Date", "Value", "Transfer Fee", "Expiry"]
        df = pd.DataFrame(columns=required_columns)  # ✅ Correct
        df.to_excel(WAHOO_VEHICLES_FILE, index=False, engine="openpyxl")


setup_wahoo_vehicles_file()

@app.route("/wahoo_vehicles")
@login_required
def wahoo_vehicles():
    try:
        vehicles = Vehicle.query.all()  # Use SQLAlchemy
        return render_template("wahoo_vehicles.html", wahoo_vehicles=vehicles)
    except Exception as e:
        flash(f"Error loading Wahoo Pool Vehicles data: {e}", "error")
        return redirect(url_for("index"))



@app.route("/add_wahoo_vehicle", methods=["POST"])
@login_required
def add_wahoo_vehicle():
    try:
        new_vehicle = Vehicle(
            plate=request.form["plate"].strip(),
            type=request.form["type"].strip(),
            vin=request.form["vin"].strip(),
            rego_renewal_date=request.form["rego_renewal_date"].strip(),
            insurance_renewal_date=request.form.get("ctp_date", "").strip(),
            value=float(request.form["value"]),
            transfer_fee=float(request.form.get("transfer_fee", 0))
        )
        db.session.add(new_vehicle)
        db.session.commit()
        flash("✅ Vehicle added successfully!", "success")
    except Exception as e:
        flash(f"Error adding vehicle: {e}", "error")
    return redirect(url_for("harm_drive"))


@app.route("/update_wahoo_vehicle", methods=["POST"])
@login_required
def update_wahoo_vehicle():
    try:
        plate_to_update = request.form["plate_to_update"].strip()
        new_rego_renewal_date = request.form["new_rego_renewal_date"].strip()

        df = pd.read_excel(WAHOO_VEHICLES_FILE, engine="openpyxl")

        if plate_to_update in df["Plate"].values:
            df.loc[df["Plate"] == plate_to_update, "Rego Renewal Date"] = new_rego_renewal_date
            df["Expiry"] = df["Rego Renewal Date"].apply(lambda x: calculate_expiry(x) if pd.notnull(x) else None)
            df.to_excel(WAHOO_VEHICLES_FILE, index=False, engine="openpyxl")
            flash(f"✅ Vehicle {plate_to_update} updated successfully!", "success")
        else:
            flash(f"⚠️ Vehicle with plate {plate_to_update} not found.", "error")
    except Exception as e:
        flash(f"⚠️ Error updating vehicle: {e}", "error")

    return redirect(url_for("wahoo_vehicles"))

@app.route("/delete_wahoo_vehicle", methods=["POST"])
@login_required
def delete_wahoo_vehicle():
    try:
        plate_to_delete = request.form["plate_to_delete"].strip()

        df = pd.read_excel(WAHOO_VEHICLES_FILE, engine="openpyxl")

        if plate_to_delete in df["Plate"].values:
            df = df[df["Plate"] != plate_to_delete]
            df.to_excel(WAHOO_VEHICLES_FILE, index=False, engine="openpyxl")
            flash(f"✅ Vehicle {plate_to_delete} deleted successfully!", "success")
        else:
            flash(f"⚠️ Vehicle with plate {plate_to_delete} not found.", "error")
    except Exception as e:
        flash(f"⚠️ Error deleting vehicle: {e}", "error")

    return redirect(url_for("wahoo_vehicles"))



@app.route('/')
def home():
    return redirect(url_for('employees'))


@app.route('/search_employee', methods=['GET', 'POST'])
@login_required
def search_employee():
    data = pd.read_excel(EMPLOYEE_FILE)  # Load data from Excel
    selected_employees = []  # List to store matched employees

    if request.method == 'POST':
        # Get search criteria from the form
        search_name = request.form.get('search_name', '').strip().lower()
        search_family_name = request.form.get('search_family_name', '').strip().lower()

        # Ensure both name and family name are provided
        if not search_name or not search_family_name:
            flash("Both Name and Family Name are required!", "error")
            return render_template('search_employee.html', selected_employees=[])

        # Normalize column names and data
        data.columns = map(str.strip, data.columns)
        if 'Name' in data.columns and 'Family Name' in data.columns:
            data['Name'] = data['Name'].astype(str).str.lower().str.strip()
            data['Family Name'] = data['Family Name'].astype(str).str.lower().str.strip()
        else:
            flash("Required columns ('Name' and 'Family Name') not found in the dataset.", "error")
            return render_template('search_employee.html', selected_employees=[])

        # Apply filtering
        filtered_data = data[
            (data['Name'] == search_name) & (data['Family Name'] == search_family_name)
        ]

        # Debugging: Print filtered data
        print("Filtered Data:\n", filtered_data)

        # Convert filtered data to dictionary
        selected_employees = filtered_data.to_dict(orient='records')

    return render_template('search_employee.html', selected_employees=selected_employees)



@app.route('/add_employee', methods=['GET', 'POST'])
def add_employee():
    if request.method == 'POST':
        # Handle adding logic
        new_employee = {
            'Name': request.form['name'],
            'Family Name': request.form['family_name'],
            'TFN': request.form['tfn'],
            'ABN': request.form['abn'],
            'Address': request.form['address'],
            'Email': request.form['email'],
            'Phone': request.form['phone']
        }
        data = pd.read_excel(EMPLOYEE_FILE)
        new_row = pd.DataFrame([new_employee])
        data = pd.concat([data, new_row], ignore_index=True)
        data.to_excel(EMPLOYEE_FILE, index=False)
        return redirect(url_for('employees'))

    return render_template('add_employee.html')

@app.route("/add_company", methods=["GET", "POST"])
@login_required
def add_company():
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

            # Create a unique folder for the company
            folder_name = secure_filename(f"{company_name}_{abn}")
            folder_path = os.path.join(MAIN_DIR, folder_name)
            os.makedirs(folder_path, exist_ok=True)

            # Handle mandatory ASIC Extract upload
            asic_extract = request.files.get("asic_extract")
            if not asic_extract or not allowed_file(asic_extract.filename):
                flash("ASIC Extract is required and must be a valid file!", "error")
                return redirect(url_for("add_company"))
            asic_extract.save(os.path.join(folder_path, secure_filename(asic_extract.filename)))

            # Handle optional document uploads
            for field_name in ["company_registration", "logo"]:
                file = request.files.get(field_name)
                if file and file.filename and allowed_file(file.filename):
                    file.save(os.path.join(folder_path, secure_filename(file.filename)))


            # Update the Excel file
            if not os.path.exists(EXCEL_FILE):
                flash("Company data file not found. Please create the file first.", "error")
                return redirect(url_for("add_company"))

            df = pd.read_excel(EXCEL_FILE, engine="openpyxl")
            if "Company Name" not in df.columns or "ABN" not in df.columns:
                flash("Missing required columns in the data file.", "error")
                return redirect(url_for("add_company"))

            if any((df["Company Name"] == company_name) & (df["ABN"] == abn)):
                flash("A company with this name and ABN already exists!", "error")
                return redirect(url_for("add_company"))

            new_row = {
                "Company Name": company_name,
                "Registration Date": registration_date,
                "ABN": abn,
                "ACN": acn,
                "Type": company_type,
                "Registered Address": registered_address,
                "QBCC License Number": qbcc_license_number,
                "Documents": folder_name,
            }
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
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

