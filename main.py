from flask import Flask, render_template, send_from_directory, request, redirect, url_for, session, flash, send_file, jsonify
import os
import pandas as pd
from werkzeug.utils import secure_filename
from datetime import datetime
import requests
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()


# Initialize Flask app
app = Flask(__name__, template_folder='Templates')
app.secret_key = "super_secret_key"

# print(f"Secret Key: {app.secret_key}")

# print(app.jinja_loader.searchpath)


# Secure Environment Variables
PASSWORD = os.getenv("ADMIN_PASSWORD", "default_fallback_password")

# Define Base Directory
BASE_DIR = os.getcwd()

# File Paths
EXCEL_FILE = os.path.join(BASE_DIR, "CompanyData.xlsx")
MAIN_DIR = os.path.join(BASE_DIR, "CompanyFolders")
EMPLOYEE_FILE = os.path.join(BASE_DIR, "employees.xlsx")
HARM_DRIVE_FILE = os.path.join(BASE_DIR, "HarmDriveData.xlsx")
PROJECTS_FILE = os.path.join(BASE_DIR, "projects.xlsx")

# Allowed File Extensions
ALLOWED_EXTENSIONS = {"pdf", "docx", "jpg", "jpeg", "png"}

# Project Images Directory
PROJECT_IMAGES_DIR = os.path.join(BASE_DIR, "static", "project_images")
os.makedirs(PROJECT_IMAGES_DIR, exist_ok=True)  # Ensure it exists

# File path for Wahoo Pool Vehicles
WAHOO_VEHICLES_FILE = "wahoo_pool_vehicles.xlsx"


if not os.path.exists(PROJECT_IMAGES_DIR):
    os.makedirs(PROJECT_IMAGES_DIR)

if not os.path.exists(PROJECTS_FILE):
    pd.DataFrame(columns=["Project Name", "Description", "Start Date", "End Date"]).to_excel(PROJECTS_FILE, index=False)


# Utility function to initialize Excel files if they don't exist
def initialize_excel_file(file_path, columns):
    """
    Initialize an Excel file with the given columns if it doesn't exist.
    :param file_path: Path to the Excel file
    :param columns: List of column names for the Excel file
    """
    try:
        if not os.path.exists(file_path):
            pd.DataFrame(columns=columns).to_excel(file_path, index=False, engine="openpyxl")
    except Exception as e:
        print(f"Error initializing file {file_path}: {e}")
        exit(1)

# Ensure environment setup
def setup_environment():
    """
    Ensure all required directories and Excel files are initialized.
    """
    try:
        # Create the main directory if it doesn't exist
        os.makedirs(MAIN_DIR, exist_ok=True)

        # Initialize required Excel files
        initialize_excel_file(
            EXCEL_FILE,
            [
                "Company Name", "Registration Date", "ABN", "ACN",
                "Type", "Registered Address", "QBCC License Number", "Documents"
            ]
        )
        initialize_excel_file(
            EMPLOYEE_FILE,
            ["Name", "Family Name", "TFN", "ABN", "Address", "Email", "Phone"]
        )
        initialize_excel_file(
            HARM_DRIVE_FILE,
            ["Plate", "Type", "VIN", "Rego Renewal Date", "Insurance Renewal (CTP) Date", "Expiry", "Value", "Transfer Fee"]
        )
        initialize_excel_file(
            PROJECTS_FILE,
            ["Project Name", "Description", "Start Date", "End Date"]
        )
    except PermissionError:
        print(f"Permission denied: Cannot create or write to required files.")
        exit(1)  # Exit gracefully if permissions are insufficient

# Utility to check allowed file extensions
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

# Middleware to protect routes
def login_required(f):
    def wrapper(*args, **kwargs):
        if not session.get("logged_in"):
            flash("You must log in to access this page.", "error")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    wrapper.__name__ = f.__name__
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
        if os.path.exists(HARM_DRIVE_FILE):
            df = pd.read_excel(HARM_DRIVE_FILE, engine="openpyxl")
            today = datetime.today()
            df["Rego Renewal Date"] = pd.to_datetime(df["Rego Renewal Date"], format="%d/%m/%Y", errors="coerce")
            df["Expiry"] = df["Rego Renewal Date"].apply(lambda x: (x - today).days if pd.notnull(x) else None)
            vehicles = df[df["Expiry"].notnull() & (df["Expiry"] <= 50)][["Plate", "Type", "Expiry"]].to_dict(orient="records")

        # Pagination Logic
        PER_PAGE = 5  # 5 news items per page
        page = int(request.args.get("page", 1))
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


@app.route("/harm_drive", methods=["GET", "POST"])
@login_required
def harm_drive():
    try:
        # Load the data from the Harm Drive Excel file or create one if it doesn't exist
        if not os.path.exists(HARM_DRIVE_FILE):
            pd.DataFrame(columns=[
                "Plate", "Type", "VIN", "Rego Renewal Date",
                "Insurance Renewal (CTP) Date", "Value", "Transfer Fee", "Expiry"
            ]).to_excel(HARM_DRIVE_FILE, index=False, engine="openpyxl")

        df = pd.read_excel(HARM_DRIVE_FILE, engine="openpyxl")

        # Ensure necessary columns exist
        required_columns = [
            "Plate", "Type", "VIN", "Rego Renewal Date",
            "Insurance Renewal (CTP) Date", "Value", "Transfer Fee", "Expiry"
        ]
        for column in required_columns:
            if column not in df.columns:
                df[column] = None

        # Format dates to DD/MM/YYYY
        def format_date_ddmmyyyy(date):
            if pd.notnull(date):
                return pd.to_datetime(date, dayfirst=True).strftime("%d/%m/%Y")
            return None

        # Calculate expiry days
        def calculate_expiry(rego_date):
            try:
                today = datetime.today()
                rego_date = pd.to_datetime(rego_date, dayfirst=True)
                return (rego_date - today).days
            except Exception:
                return None

        # Apply date formatting and expiry calculations
        df["Rego Renewal Date"] = df["Rego Renewal Date"].apply(format_date_ddmmyyyy)
        df["Insurance Renewal (CTP) Date"] = df["Insurance Renewal (CTP) Date"].apply(format_date_ddmmyyyy)
        df["Expiry"] = df["Rego Renewal Date"].apply(lambda x: calculate_expiry(x) if pd.notnull(x) else None)

        # Convert DataFrame to a list of dictionaries for rendering
        vehicles = df.to_dict(orient="records")

        # Handle POST requests for adding, updating, or deleting vehicles
        if request.method == "POST":
            action = request.form.get("action")

            if action == "add":
                try:
                    new_vehicle = {
                        "Plate": request.form["plate"].strip(),
                        "Type": request.form["type"].strip(),
                        "VIN": request.form["vin"].strip(),
                        "Rego Renewal Date": request.form["rego_renewal_date"].strip(),
                        "Insurance Renewal (CTP) Date": request.form.get("ctp_date", "").strip(),
                        "Value": float(request.form["value"]),
                        "Transfer Fee": float(request.form.get("transfer_fee", 0)),
                        "Expiry": None  # Will be recalculated
                    }
                    df = pd.concat([df, pd.DataFrame([new_vehicle])], ignore_index=True)
                    df["Rego Renewal Date"] = pd.to_datetime(df["Rego Renewal Date"], errors="coerce").dt.strftime("%d/%m/%Y")
                    df["Insurance Renewal (CTP) Date"] = pd.to_datetime(df["Insurance Renewal (CTP) Date"], errors="coerce").dt.strftime("%d/%m/%Y")
                    df["Expiry"] = df["Rego Renewal Date"].apply(
                        lambda x: calculate_expiry(x) if pd.notnull(x) else None
                    )
                    df.to_excel(HARM_DRIVE_FILE, index=False, engine="openpyxl")
                    flash("Vehicle added successfully!", "success")
                except Exception as e:
                    flash(f"Error adding vehicle: {e}", "error")

            elif action == "delete":
                try:
                    plate_to_delete = request.form["plate_to_delete"].strip()
                    if plate_to_delete in df["Plate"].values:
                        df = df[df["Plate"] != plate_to_delete]
                        df.to_excel(HARM_DRIVE_FILE, index=False, engine="openpyxl")
                        flash(f"Vehicle with plate {plate_to_delete} deleted successfully!", "success")
                    else:
                        flash(f"Vehicle with plate {plate_to_delete} not found.", "error")
                except Exception as e:
                    flash(f"Error deleting vehicle: {e}", "error")

            elif action == "update":
                try:
                    plate_to_update = request.form["plate_to_update"].strip()
                    rego_renewal = request.form["rego_renewal_update"].strip()
                    if plate_to_update in df["Plate"].values:
                        df.loc[df["Plate"] == plate_to_update, "Rego Renewal Date"] = rego_renewal
                        df["Rego Renewal Date"] = pd.to_datetime(df["Rego Renewal Date"], errors="coerce").dt.strftime("%d/%m/%Y")
                        df["Expiry"] = df["Rego Renewal Date"].apply(
                            lambda x: calculate_expiry(x) if pd.notnull(x) else None
                        )
                        df.to_excel(HARM_DRIVE_FILE, index=False, engine="openpyxl")
                        flash(f"Vehicle with plate {plate_to_update} updated successfully!", "success")
                    else:
                        flash(f"Vehicle with plate {plate_to_update} not found.", "error")
                except Exception as e:
                    flash(f"Error updating vehicle: {e}", "error")

        # Render the harm_drive template with the vehicle data
        return render_template("harm_drive.html", title="Harm Drive Pty Ltd", vehicles=vehicles)

    except Exception as e:
        flash(f"Error processing Harm Drive Pty Ltd data: {e}", "error")
        return redirect(url_for("index"))



@app.route("/update_vehicle", methods=["POST"])
def update_vehicle():
    try:
        plate_to_update = request.form["plate_to_update"].strip()
        new_rego_renewal_date = request.form["new_rego_renewal_date"].strip()

        print(f"🔹 Received update request for plate: {plate_to_update} with new date: {new_rego_renewal_date}")

        df = pd.read_excel(HARM_DRIVE_FILE, engine="openpyxl")

        if plate_to_update not in df["Plate"].values:
            print(f"❌ Plate {plate_to_update} not found in Excel.")
            flash(f"Vehicle with plate {plate_to_update} not found.", "error")
            return redirect(url_for("harm_drive"))

        print(f"✅ Plate {plate_to_update} found in Excel.")

        # Convert and format date properly
        new_rego_renewal_date = pd.to_datetime(new_rego_renewal_date, errors="coerce", dayfirst=True).strftime("%d/%m/%Y")

        # Ensure the correct row is being updated
        df.loc[df["Plate"] == plate_to_update, "Rego Renewal Date"] = new_rego_renewal_date

        # Recalculate Expiry
        today = datetime.today()
        df["Expiry"] = df["Rego Renewal Date"].apply(
            lambda x: (pd.to_datetime(x, format="%d/%m/%Y", errors="coerce") - today).days if pd.notnull(x) else None
        )

        print("✅ Updated DataFrame:")
        print(df[df["Plate"] == plate_to_update])  # Debugging output

        # Save the updated DataFrame back to the Excel file
        df.to_excel(HARM_DRIVE_FILE, index=False, engine="openpyxl")

        # Reload to confirm changes were saved
        df_check = pd.read_excel(HARM_DRIVE_FILE, engine="openpyxl")
        if plate_to_update in df_check["Plate"].values:
            saved_date = df_check.loc[df_check["Plate"] == plate_to_update, "Rego Renewal Date"].values[0]
            if saved_date == new_rego_renewal_date:
                print(f"✅ Verified: Rego Renewal Date updated successfully for {plate_to_update}")
            else:
                print(f"❌ Warning: Mismatch detected in saved date! Expected {new_rego_renewal_date}, but got {saved_date}")

        print(f"✅ Successfully saved changes to {HARM_DRIVE_FILE}")
        flash(f"Vehicle with plate {plate_to_update} updated successfully!", "success")

        return redirect(url_for("harm_drive"))

    except Exception as e:
        print(f"❌ Error updating vehicle: {e}")
        flash(f"Error updating vehicle: {e}", "error")
        return redirect(url_for("harm_drive"))



@app.route('/add_vehicle', methods=['POST'])
@login_required
def add_vehicle():
    try:
        plate = request.form['plate'].strip()
        type_ = request.form['type'].strip()
        vin = request.form['vin'].strip()
        rego_renewal_date = request.form['rego_renewal_date'].strip()
        ctp_date = request.form.get('ctp_date', '').strip()
        value = request.form['value'].strip()
        transfer_fee = request.form.get('transfer_fee', '').strip()

        new_vehicle = {
            "Plate": plate,
            "Type": type_,
            "VIN": vin,
            "Rego Renewal Date": rego_renewal_date,
            "Insurance Renewal (CTP) Date": ctp_date,
            "Expiry": None,
            "Value": value,
            "Transfer Fee": transfer_fee
        }

        df = pd.read_excel(HARM_DRIVE_FILE, engine="openpyxl")
        df = pd.concat([df, pd.DataFrame([new_vehicle])], ignore_index=True)
        df.to_excel(HARM_DRIVE_FILE, index=False, engine="openpyxl")
        flash("Vehicle added successfully!", "success")
    except Exception as e:
        flash(f"Error adding vehicle: {e}", "error")
    return redirect(url_for("harm_drive"))

@app.route('/delete_vehicle', methods=['POST'])
@login_required
def delete_vehicle():
    try:
        plate = request.form['plate_to_delete'].strip()

        df = pd.read_excel(HARM_DRIVE_FILE, engine="openpyxl")
        df = df[df['Plate'] != plate]
        df.to_excel(HARM_DRIVE_FILE, index=False, engine="openpyxl")
        flash(f"Vehicle with plate {plate} deleted successfully!", "success")
    except Exception as e:
        flash(f"Error deleting vehicle: {e}", "error")
    return redirect(url_for("harm_drive"))


# Ensure the Wahoo Vehicles file exists
def setup_wahoo_vehicles_file():
    if not os.path.exists(WAHOO_VEHICLES_FILE):
        columns = ["Plate", "Type", "Expiry Date"]
        df = pd.DataFrame(columns=columns)
        df.to_excel(WAHOO_VEHICLES_FILE, index=False, engine="openpyxl")


setup_wahoo_vehicles_file()

# ✅ Wahoo Pool Vehicles Routes
@app.route("/wahoo_vehicles")
@login_required
def wahoo_vehicles():
    try:
        df = pd.read_excel(WAHOO_VEHICLES_FILE, engine="openpyxl")

        # Ensure required columns exist
        required_columns = ["Plate", "Type", "VIN", "Rego Renewal Date", "Insurance Renewal (CTP) Date", "Value", "Transfer Fee", "Expiry"]
        for column in required_columns:
            if column not in df.columns:
                df[column] = None

        df["Rego Renewal Date"] = pd.to_datetime(df["Rego Renewal Date"], errors="coerce")  # Keep as datetime
        df["Expiry"] = df["Rego Renewal Date"].apply(lambda x: calculate_expiry(x) if pd.notnull(x) else None)
        df["Rego Renewal Date"] = df["Rego Renewal Date"].dt.strftime("%Y-%m-%d")  # Convert to string only for saving


        wahoo_vehicles_list = df.to_dict(orient="records")
    except Exception as e:
        flash(f"Error loading Wahoo Pool Vehicles data: {e}", "error")
        wahoo_vehicles_list = []

    return render_template("wahoo_vehicles.html", wahoo_vehicles=wahoo_vehicles_list)


@app.route("/add_wahoo_vehicle", methods=["POST"])
@login_required
def add_wahoo_vehicle():
    try:
        plate = request.form["plate"].strip()
        type_ = request.form["type"].strip()
        vin = request.form["vin"].strip()
        rego_renewal_date = request.form["rego_renewal_date"].strip()
        ctp_date = request.form.get("ctp_date", "").strip()
        value = request.form["value"].strip()
        transfer_fee = request.form.get("transfer_fee", "").strip()

        df = pd.read_excel(WAHOO_VEHICLES_FILE, engine="openpyxl")

        new_vehicle = {
            "Plate": plate,
            "Type": type_,
            "VIN": vin,
            "Rego Renewal Date": rego_renewal_date,
            "Insurance Renewal (CTP) Date": ctp_date,
            "Expiry": calculate_expiry(rego_renewal_date),
            "Value": value,
            "Transfer Fee": transfer_fee
        }

        df = pd.concat([df, pd.DataFrame([new_vehicle])], ignore_index=True)
        df.to_excel(WAHOO_VEHICLES_FILE, index=False, engine="openpyxl")

        flash("✅ Vehicle added successfully!", "success")
    except Exception as e:
        flash(f"⚠️ Error adding vehicle: {e}", "error")

    return redirect(url_for("wahoo_vehicles"))

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
                if file and allowed_file(file.filename):
                    saved_path = os.path.join(folder_path, secure_filename(file.filename))
                    file.save(saved_path)
                    app.logger.info(f"Saved {field_name} to {saved_path}")

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




@app.route("/download/<path:company_folder>/<path:filename>")
@login_required
def download_file(company_folder, filename):
    try:
        # Construct the full file path inside the correct company folder
        filepath = os.path.join(MAIN_DIR, company_folder, filename)

        if not os.path.exists(filepath):
            flash("File not found!", "error")
            return redirect(url_for("view_company"))

        return send_file(filepath, as_attachment=True)
    except Exception as e:
        flash(f"Error downloading file: {e}", "error")
        return redirect(url_for("view_company"))



if __name__ == "__main__":
    app.run(debug=True)
