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
# "Center" is included at the top of the list
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
            session['login_time'] = datetime.datetime.now().strftime("%d-%m-%Y %H:%M")
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
    
    # Calculate dynamic greeting
    hour = datetime.datetime.now().hour
    if hour < 12:
        greeting = "Good Morning"
    elif hour < 17:
        greeting = "Good Afternoon"
    else:
        greeting = "Good Evening"
    
    return render_template('dashboard.html', 
                           total=total_records, 
                           circle_count=circle_records, 
                           table=data_html,
                           greeting=greeting)

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
