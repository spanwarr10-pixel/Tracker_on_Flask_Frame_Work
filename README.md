# Tracker_on_Flask_Frame_Work
Flask is a UI based frame work on this we deploy Project Tracker.
1. Project Directory Structure
wow_flask_app/
│
├── app.py                  # Main Python backend script
├── requirements.txt        # Python dependencies
│
└── templates/              # HTML Frontend files
    ├── base.html           # Master layout (Sidebar & CSS theme)
    ├── login.html          # Login page
    ├── dashboard.html      # Main Dashboard
    ├── upload.html         # Upload Data Hub
    └── summary.html        # Other Summary (Logs)


Required Libraries (requirements.txt)- install them using pip install -r requirements.txt
Flask==3.0.0
pandas==2.1.1
openpyxl==3.1.2
Werkzeug==3.0.0

The Backend (app.py) - Code
**#Code Start**
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
import pandas as pd
import sqlite3
import datetime
import io
import os
from functools import wraps

app = Flask(__name__)
app.secret_key = 'airtel_priority_secret_key_change_in_production'

DB_PATH = "wow_tracker_master.db"
COLUMNS_SCHEMA = [
    "WoW", "Circle", "Actioned Month", "Claimed Date", "Project", 
    "Project Category", "# Site ID", "TOCO Name", "TOCO ID", "IP Fee", 
    "LLR", "Loading", "EB", "Diesel", "Claimed WoW", "New Site Id", "New ToCo id"
]
CIRCLES = ["Center", "AP", "ASM", "BIH", "DEL", "GUJ", "HP", "HRY", "JK", "KER", "KK", 
           "KOL", "MAH", "MP", "MUM", "ORI", "PUN", "RAJ", "TN", "UPE", "UPW", "WB"]

# --- DATABASE SETUP ---
def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    cols_def = ", ".join([f'"{col}" TEXT' for col in COLUMNS_SCHEMA])
    c.execute(f'CREATE TABLE IF NOT EXISTS wow_data ({cols_def})')
    c.execute('''CREATE TABLE IF NOT EXISTS upload_logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, user_name TEXT, olm_id TEXT, 
                    circle TEXT, upload_date DATE, upload_time TEXT, records_added INTEGER)''')
    conn.commit()
    conn.close()

init_db()

# --- AUTHENTICATION DECORATOR ---
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'logged_in' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

# --- VALIDATION LOGIC ---
def validate_uploaded_data(df):
    missing_cols = [col for col in COLUMNS_SCHEMA if col not in df.columns]
    if missing_cols:
        return False, [f"Missing required columns: {', '.join(missing_cols)}"], None

    df = df[COLUMNS_SCHEMA].copy()
    df_str = df.fillna("").astype(str)
    errors = []
    
    for index, row in df_str.iterrows():
        wow_type = row['WoW'].strip().lower()
        if wow_type not in ['claimed', 'funnel']:
            errors.append(f"Row {index + 2}: 'WoW' column must be 'Claimed' or 'Funnel'.")
            continue
            
        for col in COLUMNS_SCHEMA:
            val = row[col].strip()
            if wow_type == 'claimed' and val == "":
                errors.append(f"Row {index + 2}: '{col}' is mandatory when WoW is 'Claimed'.")
            elif wow_type == 'funnel' and val == "" and col != 'Claimed Date':
                errors.append(f"Row {index + 2}: '{col}' is mandatory when WoW is 'Funnel'.")
                    
    if errors: return False, errors, None
    return True, "Success", df

# --- ROUTES ---
@app.route('/', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form.get('email')
        olm = request.form.get('olm')
        circle = request.form.get('circle')
        
        if email and olm and circle and "@airtel.com" in email.lower():
            session['logged_in'] = True
            session['email'] = email
            session['olm'] = olm
            session['circle'] = circle
            return redirect(url_for('dashboard'))
        else:
            flash('Invalid login. Ensure you use an @airtel.com email.', 'danger')
    return render_template('login.html', circles=CIRCLES)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/dashboard')
@login_required
def dashboard():
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql("SELECT * FROM wow_data", conn)
    conn.close()
    
    total_records = len(df)
    circle_records = len(df[df['Circle'] == session.get('circle')])
    data_html = df.to_html(classes="table table-striped table-bordered", index=False) if not df.empty else None
    
    return render_template('dashboard.html', total=total_records, circle_count=circle_records, table=data_html)

@app.route('/upload', methods=['GET', 'POST'])
@login_required
def upload():
    errors = []
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part', 'danger')
            return redirect(request.url)
        
        file = request.files['file']
        if file.filename == '':
            flash('No selected file', 'danger')
            return redirect(request.url)

        if file and (file.filename.endswith('.csv') or file.filename.endswith('.xlsx')):
            try:
                if file.filename.endswith('.csv'):
                    raw_df = pd.read_csv(file)
                else:
                    raw_df = pd.read_excel(file)
                
                # Remove duplicates exactly like requested earlier
                raw_df = raw_df.drop_duplicates(ignore_index=True)
                is_valid, message, final_df = validate_uploaded_data(raw_df)
                
                if is_valid:
                    conn = sqlite3.connect(DB_PATH)
                    final_df.to_sql('wow_data', conn, if_exists='append', index=False)
                    now = datetime.datetime.now()
                    c = conn.cursor()
                    c.execute('INSERT INTO upload_logs (user_name, olm_id, circle, upload_date, upload_time, records_added) VALUES (?, ?, ?, ?, ?, ?)', 
                              (session['email'], session['olm'], session['circle'], now.strftime("%Y-%m-%d"), now.strftime("%H:%M:%S"), len(final_df)))
                    conn.commit()
                    conn.close()
                    flash(f'Successfully uploaded {len(final_df)} records!', 'success')
                else:
                    errors = message
            except Exception as e:
                flash(f'Error processing file: {str(e)}', 'danger')
                
    return render_template('upload.html', errors=errors)

@app.route('/download_template')
@login_required
def download_template():
    template_df = pd.DataFrame(columns=COLUMNS_SCHEMA)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        template_df.to_excel(writer, index=False)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name="WOW_Project_Template.xlsx", mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/summary')
@login_required
def summary():
    today_str = datetime.datetime.now().strftime("%Y-%m-%d")
    conn = sqlite3.connect(DB_PATH)
    query = f"SELECT user_name as Name, olm_id as OLM, circle as Circle, upload_time as Time, records_added as Records FROM upload_logs WHERE upload_date = '{today_str}' ORDER BY Time DESC"
    df = pd.read_sql(query, conn)
    conn.close()
    
    data_html = df.to_html(classes="table table-striped table-bordered", index=False) if not df.empty else None
    return render_template('summary.html', table=data_html)

if __name__ == '__main__':
    app.run(debug=True)
    **#Code end**


**Frontend Templates (templates/ folder)**
base.html (The Master UI Layout)
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>WOW Project Dashboard</title>
    <style>
        body { margin: 0; font-family: 'Arial', sans-serif; background-color: #F8F9FA; display: flex; height: 100vh; }
        
        /* Sidebar Styles (Black & Red) */
        .sidebar { width: 250px; background-color: #000000; color: #fff; display: flex; flex-direction: column; border-right: 2px solid #FF0000; }
        .sidebar-header { text-align: center; padding: 20px 0; border-bottom: 1px solid #333; }
        .sidebar-header h1 { color: #FF0000; margin: 0; font-size: 28px; }
        .sidebar-header p { color: #aaa; font-size: 12px; margin: 5px 0 0 0; }
        .nav-links { padding: 20px 10px; flex-grow: 1; }
        .nav-links a { display: block; padding: 12px; color: #fff; text-decoration: none; border: 1px solid #333; border-radius: 4px; margin-bottom: 10px; font-weight: bold; text-align: center; transition: 0.3s; }
        .nav-links a:hover { border-color: #FF0000; color: #FF0000; background-color: rgba(255,0,0,0.1); }
        .sidebar-footer { text-align: center; padding: 20px; font-size: 12px; color: #666; }
        
        /* Main Content Styles */
        .main-content { flex-grow: 1; padding: 30px; overflow-y: auto; }
        h2 { color: #000; margin-top: 0; }
        .airtel-red { color: #FF0000; }
        
        /* Alerts & Tables */
        .alert { padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        .alert-success { background-color: #D4EDDA; color: #155724; border: 1px solid #C3E6CB; }
        .alert-danger { background-color: #F8D7DA; color: #721C24; border: 1px solid #F5C6CB; }
        .table { width: 100%; border-collapse: collapse; margin-top: 20px; background: #fff; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
        .table th, .table td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        .table th { background-color: #f2f2f2; }
    </style>
</head>
<body>
    {% if session.get('logged_in') %}
    <div class="sidebar">
        <div class="sidebar-header">
            <h1>airtel</h1>
            <p>User: {{ session.olm }} | {{ session.circle }}</p>
        </div>
        <div class="nav-links">
            <a href="{{ url_for('dashboard') }}">📊 Main Dashboard</a>
            <a href="{{ url_for('upload') }}">📤 Upload</a>
            <a href="{{ url_for('summary') }}">📋 Other Summary</a>
        </div>
        <div class="nav-links" style="flex-grow: 0;">
            <a href="{{ url_for('logout') }}" style="border-color: #666; color: #aaa;">🚪 Logout</a>
        </div>
        <div class="sidebar-footer">
            Support: Sandeep Panwar<br>Sandeep.Panwar@airtel.com
        </div>
    </div>
    {% endif %}

    <div class="main-content">
        {% with messages = get_flashed_messages(with_categories=true) %}
          {% if messages %}
            {% for category, message in messages %}
              <div class="alert alert-{{ category }}">{{ message }}</div>
            {% endfor %}
          {% endif %}
        {% endwith %}
        
        {% block content %}{% endblock %}
    </div>
</body>
</html>


**login.html**
{% extends "base.html" %}
{% block content %}
<div style="max-width: 400px; margin: 100px auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 4px 10px rgba(0,0,0,0.1);">
    <h1 style="text-align: center;"><span class="airtel-red">airtel</span> priority</h1>
    <h3 style="text-align: center; color: #555; margin-bottom: 30px;">WOW Project Tracker</h3>
    
    <form method="POST" action="{{ url_for('login') }}">
        <div style="margin-bottom: 15px;">
            <label style="display: block; margin-bottom: 5px;">Airtel Email ID</label>
            <input type="email" name="email" required placeholder="user@airtel.com" style="width: 100%; padding: 10px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box;">
        </div>
        <div style="margin-bottom: 15px;">
            <label style="display: block; margin-bottom: 5px;">OLM ID</label>
            <input type="text" name="olm" required placeholder="Enter OLM ID" style="width: 100%; padding: 10px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box;">
        </div>
        <div style="margin-bottom: 25px;">
            <label style="display: block; margin-bottom: 5px;">Select Circle</label>
            <select name="circle" required style="width: 100%; padding: 10px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box;">
                <option value="">Select...</option>
                {% for c in circles %}
                <option value="{{ c }}">{{ c }}</option>
                {% endfor %}
            </select>
        </div>
        <button type="submit" style="width: 100%; padding: 12px; background-color: #FF0000; color: white; border: none; border-radius: 4px; font-weight: bold; cursor: pointer;">Access Workspace</button>
    </form>
</div>
{% endblock %}


**dashboard.html**
{% extends "base.html" %}
{% block content %}
<h2><span class="airtel-red">WOW Project</span> Dashboard</h2>

<div style="display: flex; gap: 20px; margin-bottom: 20px;">
    <div style="flex: 1; background: #E2E8F0; padding: 15px; border-radius: 5px;">
        <strong>Total Records Uploaded:</strong> {{ total }}
    </div>
    <div style="flex: 1; background: #D4EDDA; padding: 15px; border-radius: 5px;">
        <strong>Records for {{ session.circle }} Circle:</strong> {{ circle_count }}
    </div>
</div>

<h3>📊 Database View</h3>
<div style="overflow-x: auto;">
    {% if table %}
        {{ table|safe }}
    {% else %}
        <p>No data available yet. Please navigate to the Upload page.</p>
    {% endif %}
</div>
{% endblock %}

**upload.html**
{% extends "base.html" %}
{% block content %}
<h2><span class="airtel-red">Upload</span> Data Hub</h2>
<hr>

<div style="display: flex; gap: 40px; margin-top: 20px;">
    <div style="flex: 6;">
        <h3>📤 Upload File (Excel / CSV)</h3>
        <form method="POST" enctype="multipart/form-data" style="border: 2px dashed #FF0000; padding: 30px; border-radius: 8px; text-align: center; background: #fff;">
            <input type="file" name="file" accept=".csv, .xlsx" required style="margin-bottom: 20px;">
            <br>
            <button type="submit" style="padding: 10px 20px; background-color: #FF0000; color: white; border: none; border-radius: 4px; font-weight: bold; cursor: pointer;">Validate & Process Data</button>
        </form>

        {% if errors %}
            <div class="alert alert-danger" style="margin-top: 20px;">
                <strong>Validation Failed:</strong>
                <ul style="margin-top: 10px; margin-bottom: 0;">
                {% for error in errors %}
                    <li>{{ error }}</li>
                {% endfor %}
                </ul>
            </div>
        {% endif %}
    </div>

    <div style="flex: 4;">
        <h3>📥 Download Template</h3>
        <p style="color: #666;">Use this exact format to ensure smooth uploads.</p>
        <a href="{{ url_for('download_template') }}" style="display: inline-block; padding: 10px 20px; border: 1px solid #333; color: #000; text-decoration: none; border-radius: 4px; font-weight: bold;">Download Master Template (.xlsx)</a>
    </div>
</div>
{% endblock %}

**summary.html**
{% extends "base.html" %}
{% block content %}
<h2><span class="airtel-red">Other</span> Summary</h2>
<h3>Today's Upload Log</h3>

<div style="overflow-x: auto;">
    {% if table %}
        {{ table|safe }}
    {% else %}
        <p>No data has been uploaded today.</p>
    {% endif %}
</div>
{% endblock %}



Final
To Run the Flask App:
Ensure your files are arranged in the folder structure.
Open a terminal in the wow_flask_app directory.
Run python app.py
Open a web browser and go to http://127.0.0.1:5000 or which evershowing in terminal  - I used Anaconda power shell terminal. 
