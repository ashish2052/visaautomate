
import streamlit as st
import pandas as pd
from datetime import datetime, time
import json
import os
from urllib.parse import quote

# REPORT 5: ATTENDANCE & WARNING AUTOMATION
st.set_page_config(page_title="Attendance Report", page_icon="📅", layout="wide")

# CONSTANTS
LATE_THRESHOLD = time(9, 45) # Late after 9:45
REQUIRED_HOURS = 8.0 # Minimum hours
EXIT_THRESHOLD = time(17, 30) # Early exit before 5:30

def load_config():
    if os.path.exists('config.json'):
        with open('config.json', 'r') as f:
            return json.load(f)
    return {}

def process_attendance_simple(df):
    df.columns = df.columns.str.strip()
    
    name_col = None
    datetime_col = None
    
    for col in df.columns:
        col_lower = col.lower()
        if 'name' in col_lower and 'department' not in col_lower:
            name_col = col
        elif 'date/time' in col_lower or 'datetime' in col_lower or ('date' in col_lower and 'time' in col_lower):
            datetime_col = col
    
    if not name_col or not datetime_col:
        st.error("⚠️ Could not find Name and Date/Time columns")
        return None
    
    df['Timestamp'] = pd.to_datetime(df[datetime_col], dayfirst=True, errors='coerce')
    df = df.dropna(subset=['Timestamp'])
    df['Date'] = df['Timestamp'].dt.strftime('%Y-%m-%d')
    df['Employee'] = df[name_col].astype(str).str.strip()
    
    results = []
    grouped = df.groupby(['Employee', 'Date'])
    
    for (emp, date), group in grouped:
        punches = group['Timestamp'].sort_values()
        
        if len(punches) < 1:
            continue
        
        first_in = punches.iloc[0]
        last_out = punches.iloc[-1]
        work_hours = (last_out - first_in).total_seconds() / 3600
        
        is_late = first_in.time() > LATE_THRESHOLD
        
        if work_hours >= REQUIRED_HOURS:
            is_early_exit = False
        else:
            is_early_exit = last_out.time() < EXIT_THRESHOLD
        
        is_compliant = work_hours >= REQUIRED_HOURS and not is_late
        
        note = "Compliant"
        if is_late and is_early_exit:
            note = "Late Entry & Early Exit"
        elif is_late:
            note = "Late Entry"
        elif is_early_exit:
            note = "Early Exit"
        
        results.append({
            'Employee': emp,
            'Date': date,
            'FirstIn': first_in.strftime('%H:%M:%S'),
            'LastOut': last_out.strftime('%H:%M:%S'),
            'WorkHours': round(work_hours, 1),
            'IsLate': is_late,
            'IsEarlyExit': is_early_exit,
            'IsCompliant': is_compliant,
            'Note': note
        })
    
    return pd.DataFrame(results)

HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
    /* RESET & BASE STYLES */
    :root {
        --bg-color: #1a1f2e;
        --card-bg: #232d3f;
        --text-color: #ffffff;
        --text-muted: #95a5a6;
        --green: #27ae60;
        --blue: #2980b9;
        --red: #c0392b; 
        --orange: #d35400;
        --border-color: #34495e;
    }
    
    * { box-sizing: border-box; }

    body {
        font-family: 'Inter', sans-serif;
        background-color: var(--bg-color);
        color: var(--text-color);
        margin: 0;
        padding: 20px;
        overflow-x: hidden;
    }
    
    /* HEADER WITH ACTIONS */
    .header {
        background: linear-gradient(90deg, #2c3e50 0%, #3d5a80 100%);
        padding: 15px 20px;
        border-radius: 6px;
        margin-bottom: 10px;
        border-bottom: 2px solid #3498db;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .header h1 { margin: 0; font-size: 20px; font-weight: 600; letter-spacing: 0.5px; }
    
    .action-btn {
        background: #e74c3c;
        color: white;
        border: none;
        padding: 8px 15px;
        border-radius: 4px;
        font-size: 13px;
        font-weight: 600;
        cursor: pointer;
        display: flex;
        align-items: center;
        gap: 6px;
        transition: background 0.2s;
        box-shadow: 0 2px 4px rgba(0,0,0,0.2);
    }
    .action-btn:hover { background: #c0392b; transform: translateY(-1px); }
    
    /* MODAL */
    .modal { display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; overflow: auto; background-color: rgba(0,0,0,0.7); backdrop-filter: blur(4px); }
    .modal-content { background-color: #2c3e50; margin: 5% auto; padding: 0; border: 1px solid #888; width: 80%; max-width: 900px; border-radius: 8px; box-shadow: 0 4px 20px rgba(0,0,0,0.5); animation: slideDown 0.3s ease-out; }
    @keyframes slideDown { from {transform: translateY(-50px); opacity: 0;} to {transform: translateY(0); opacity: 1;} }
    .close { color: #aaa; float: right; font-size: 28px; font-weight: bold; margin-right: 15px; cursor: pointer; }
    .close:hover { color: #fff; text-decoration: none; }
    .modal-header { padding: 15px 20px; background: #34495e; border-bottom: 1px solid #465c71; display: flex; justify-content: space-between; align-items: center; border-radius: 8px 8px 0 0; }
    .modal-body { padding: 20px; max-height: 70vh; overflow-y: auto; color: #ecf0f1; }
    
    /* DRAFT STYLE */
    .draft-item { background: white; color: #333; padding: 20px; margin-bottom: 20px; border-radius: 4px; }
    .draft-subject { font-weight: bold; margin-bottom: 10px; border-bottom: 1px solid #eee; padding-bottom: 5px; }
    
    /* TABLES */
    .data-table { width: 100%; border-collapse: collapse; font-size: 13px; }
    .data-table th { text-align: left; padding: 10px; background: rgba(255,255,255,0.05); color: #ccc; }
    .data-table td { padding: 10px; border-bottom: 1px solid rgba(255,255,255,0.05); color: #eee; }
    .emp-check { width: 16px; height: 16px; cursor: pointer; }

</style>
<script>
    function openModal() { document.getElementById('warningModal').style.display = 'block'; }
    function closeModal() { document.getElementById('warningModal').style.display = 'none'; }
    
    // Selecting
    function toggleSelect(source) {
        const checkboxes = document.querySelectorAll('.emp-check');
        for(let i=0; i<checkboxes.length; i++) {
            checkboxes[i].checked = source.checked;
        }
    }

    function generateEmails() {
        const checks = document.querySelectorAll('.emp-check:checked');
        if (checks.length === 0) { alert("Please select at least one employee."); return; }
        
        let htmlContent = "";
        
        checks.forEach(chk => {
            const empName = chk.value;
            const violations = JSON.parse(decodeURIComponent(chk.getAttribute('data-violations')));
            
            // Build Violation Table
            let table = `<table style="border-collapse: collapse; width: 100%; border: 1px solid #ccc; font-family: sans-serif; font-size: 13px; margin: 10px 0;">
                <tr style="background: #f8f9fa;">
                    <th style="border: 1px solid #ccc; padding: 8px;">Date</th>
                    <th style="border: 1px solid #ccc; padding: 8px;">In</th>
                    <th style="border: 1px solid #ccc; padding: 8px;">Out</th>
                    <th style="border: 1px solid #ccc; padding: 8px;">Hrs</th>
                    <th style="border: 1px solid #ccc; padding: 8px;">Issue</th>
                </tr>`;
            
            violations.forEach(v => {
                let color = v.IsLate ? "color:#d35400;" : (v.IsEarlyExit ? "color:#c0392b;" : "");
                table += `<tr>
                    <td style="border: 1px solid #ccc; padding: 8px;">${v.Date}</td>
                    <td style="border: 1px solid #ccc; padding: 8px; ${v.IsLate ? 'color:red;' : ''}">${v.FirstIn}</td>
                    <td style="border: 1px solid #ccc; padding: 8px; ${v.IsEarlyExit ? 'color:red;' : ''}">${v.LastOut}</td>
                    <td style="border: 1px solid #ccc; padding: 8px; font-weight:bold;">${v.WorkHours}</td>
                    <td style="border: 1px solid #ccc; padding: 8px; ${color}">${v.Note}</td>
                </tr>`;
            });
            table += "</table>";
            
            htmlContent += `
            <div class="draft-item">
                <div class="draft-subject">Subject: Notice of Attendance Irregularity - ${empName}</div>
                <div>
                    Dear ${empName},<br><br>
                    We have noticed some irregularities in your attendance this month.<br>
                    Our office hours are from <b>9:30 AM to 5:30 PM</b>.<br><br>
                    Below is a summary of dates where you were flagged:<br>
                    ${table}
                    <br>
                    Please ensure you adhere to the schedule.<br>
                    Regards,<br>Management
                </div>
            </div>`;
        });
        
        document.getElementById('modalBody').innerHTML = htmlContent;
        openModal();
    }
</script>
</head>
<body>

    <div class="header">
        <h1>Monthly Management Snapshot</h1>
        <button class="action-btn" onclick="generateEmails()">
            📩 Generate Warning Drafts
        </button>
    </div>

    <!-- METRICS PLACEHOLDER -->
    <div id="metrics"></div>

    <!-- MAIN CONTENT -->
    <div class="section-container" style="margin-top: 20px;">
        <div class="section-header">
            Risk Candidates (Late/Early/Short Hours)
            <div style="float:right;">
                <label><input type="checkbox" onclick="toggleSelect(this)"> Select All</label>
            </div>
        </div>
        <div id="riskTable"></div>
    </div>

    <!-- MODAL -->
    <div id="warningModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h2 style="margin:0; font-size:18px; color:white;">Generated Email Drafts</h2>
                <span class="close" onclick="closeModal()">&times;</span>
            </div>
            <div id="modalBody" class="modal-body">
                <!-- Drafts go here -->
            </div>
            <div style="padding:15px; background:#34495e; text-align:right;">
                <button class="action-btn" onclick="closeModal()">Close</button>
            </div>
        </div>
    </div>

</body>
</html>
"""

st.markdown("### 📊 Monthly Attendance & Warning Automation")

uploaded_file = st.file_uploader("Upload Monthly Attendance Excel", type=['xlsx', 'xls'])

if uploaded_file:
    try:
        # Determine engine based on extension
        file_ext = uploaded_file.name.split('.')[-1].lower()
        if file_ext == 'xls':
            engine = 'xlrd'
        else:
            engine = 'openpyxl'
            
        df = pd.read_excel(uploaded_file, header=0, engine=engine)
        df_daily = process_attendance_simple(df)
        
        if df_daily is not None:
            # Filter for Risk Candidates (Late, Early, or < 8 hours)
            risk_df = df_daily[
                (df_daily['IsLate']) | 
                (df_daily['IsEarlyExit']) | 
                (df_daily['WorkHours'] < REQUIRED_HOURS)
            ].copy()
            
            # Group by Employee to count violations
            stats = risk_df.groupby('Employee').size().reset_index(name='Violations')
            
            # HTML Table Generation
            table_rows = ""
            if not risk_df.empty:
                # Group by Employee for the details JSON
                for emp in stats['Employee'].unique():
                    emp_violations = risk_df[risk_df['Employee'] == emp]
                    # Convert to JSON for the data attribute
                    json_data = emp_violations.to_json(orient='records', date_format='iso')
                    clean_json = quote(json_data) # URL Encode to be safe in attribute
                    
                    count = len(emp_violations)
                    last_violation = emp_violations['Date'].max()
                    
                    table_rows += f"""
                    <tr>
                        <td>
                            <input type="checkbox" class="emp-check" value="{emp}" data-violations="{clean_json}">
                        </td>
                        <td style="color:#3498db; font-weight:500;">{emp}</td>
                        <td style="text-align:center;"><span style="background:#c0392b; padding:2px 8px; border-radius:10px; font-size:11px;">{count} Issues</span></td>
                        <td>{last_violation}</td>
                    </tr>
                    """
            else:
                table_rows = "<tr><td colspan='4' style='text-align:center; padding:20px;'>No irregularities found!</td></tr>"

            # Inject Content into Template
            # We inject the table rows into the empty div we left
            full_html = HTML_TEMPLATE.replace(
                '<div id="riskTable"></div>', 
                f"""<table class="data-table">
                    <tr>
                        <th style="width:40px;"></th>
                        <th>Employee</th>
                        <th style="text-align:center;">Violations</th>
                        <th>Last Issue</th>
                    </tr>
                    {table_rows}
                </table>"""
            )
            
            # Calculate Metrics
            total_emps = df_daily['Employee'].nunique()
            total_late = df_daily['IsLate'].sum()
            total_early = df_daily['IsEarlyExit'].sum()
            low_hours = (df_daily['WorkHours'] < REQUIRED_HOURS).sum()
            
            metrics_html = f"""
            <div class="metrics-row" style="display:grid; grid-template-columns: repeat(4, 1fr); gap:15px; margin-bottom:20px;">
                <div class="metric-card card-blue"><div class="metric-title">Active Staff</div><div class="metric-value">{total_emps}</div></div>
                <div class="metric-card card-orange"><div class="metric-title">Late Arrivals</div><div class="metric-value">{total_late}</div></div>
                <div class="metric-card card-red"><div class="metric-title">Early Exits</div><div class="metric-value">{total_early}</div></div>
                <div class="metric-card card-red"><div class="metric-title">Under 8 Hours</div><div class="metric-value">{low_hours}</div></div>
            </div>
            """
            
            full_html = full_html.replace('<div id="metrics"></div>', metrics_html)
            
            st.components.v1.html(full_html, height=850, scrolling=True)
            
    except Exception as e:
        st.error(f"Error processing file: {e}")

# IMPORTANT: No python admin section anymore.
