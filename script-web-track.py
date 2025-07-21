# app.py - Main Flask application
from flask import Flask, render_template, request, jsonify, session, send_file
from flask_cors import CORS
import pandas as pd
import sys
import os
import webbrowser
import threading
import time
import uuid
from werkzeug.utils import secure_filename
import tempfile
import json
from datetime import datetime
import io
import openpyxl
import paramiko
import re
import socket
import  dotenv
from dotenv import load_dotenv

load_dotenv()

# Use os.path.join for better cross-platform compatibility
app = Flask(__name__, template_folder='templates')
CORS(app)  # Enable CORS for all routes
app.secret_key = os.getenv('FLASK_SECRET_KEY', 'your-secret-key-change-this')  # Change this in production
app.config['MAX_CONTENT_LENGTH'] = int(os.getenv('MAX_CONTENT_LENGTH', 52428800))  # 50MB max file size

# Global storage for workbook data (in production, use Redis or database)
workbook_storage = {}

# MikroTik connection settings
MIKROTIK_SSH_PORT = int(os.getenv('MIKROTIK_SSH_PORT'))
MIKROTIK_USERNAME = os.getenv('MIKROTIK_USERNAME')
MIKROTIK_PASSWORD = os.getenv('MIKROTIK_PASSWORD')

def show_startup_message():
    """Show startup message"""
    if getattr(sys, 'frozen', False):
        print("\n" + "=" * 60)
        print("           WEB TRACK CSID - PORTABLE VERSION")
        print("=" * 60)
        print("Starting the application...")
        print("Browser will open automatically in 3 seconds...")
        print("\nIf browser doesn't open, manually go to:")
        print("http://localhost:5050")
        print("\nTo stop the application, close this window.")
        print("=" * 60 + "\n")

def open_browser():
    """Open browser after Flask starts"""
    time.sleep(3)
    try:
        webbrowser.open('http://localhost:5050')
        print("✓ Browser opened successfully!")
    except Exception as e:
        print(f"✗ Could not open browser automatically: {e}")
        print("Please manually open: http://localhost:5050")

def load_devices_active():
    """
    Load devices from JSON file
    """
    try:
        # Check if devices.json exists in the same directory as app.py
        devices_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'devices-active.json')
        if os.path.exists(devices_file):
            with open(devices_file, 'r') as f:
                devices = json.load(f)
            return devices
        else:
            # Return empty list if file doesn't exist
            print("Warning: devices.json not found. Using empty device list.")
            return []
    except Exception as e:
        print(f"Error loading devices: {e}")
        return []

def load_devices_suspend():
    """
    Load devices from JSON file
    """
    try:
        # Check if devices.json exists in the same directory as app.py
        devices_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'devices-suspend.json')
        if os.path.exists(devices_file):
            with open(devices_file, 'r') as f:
                devices = json.load(f)
            return devices
        else:
            # Return empty list if file doesn't exist
            print("Warning: devices.json not found. Using empty device list.")
            return []
    except Exception as e:
        print(f"Error loading devices: {e}")
        return []

def connect_to_mikrotik(ip, command):
    """
    Enhanced MikroTik connection function with better error handling
    """
    ssh_client = None
    try:
        # Create SSH client with more specific settings
        ssh_client = paramiko.SSHClient()
        ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        
        # Additional SSH client configuration
        ssh_client.load_system_host_keys()
        
        print(f"Attempting to connect to {ip}:{MIKROTIK_SSH_PORT}")
        
        # Connect to MikroTik with enhanced parameters
        ssh_client.connect(
            hostname=ip,
            username=MIKROTIK_USERNAME,
            password=MIKROTIK_PASSWORD,
            port=MIKROTIK_SSH_PORT,
            timeout=30,  # Increased timeout
            auth_timeout=30,  # Authentication timeout
            banner_timeout=30,  # Banner timeout
            look_for_keys=False,  # Don't look for SSH keys
            allow_agent=False,  # Don't use SSH agent
            compress=False  # Disable compression
        )
        
        print(f"Successfully connected to {ip}")
        
        # Execute command with explicit channel handling
        stdin, stdout, stderr = ssh_client.exec_command(command, timeout=30)
        
        # Wait for command to complete
        exit_status = stdout.channel.recv_exit_status()
        
        # Get output
        output = stdout.read().decode('utf-8', errors='ignore')
        error = stderr.read().decode('utf-8', errors='ignore')
        
        print(f"Command executed. Exit status: {exit_status}")
        print(f"Output length: {len(output)} characters")
        
        if exit_status != 0:
            return False, f"Command failed with exit status {exit_status}: {error}"
        
        if error and "syntax error" in error.lower():
            return False, f"MikroTik syntax error: {error}"
        
        return True, output
        
    except paramiko.AuthenticationException as e:
        error_msg = f"Authentication failed for {ip}: {str(e)}"
        print(error_msg)
        return False, error_msg
        
    except paramiko.SSHException as e:
        error_msg = f"SSH connection failed to {ip}: {str(e)}"
        print(error_msg)
        return False, error_msg
        
    except socket.timeout:
        error_msg = f"Connection timeout to {ip}:{SSH_PORT}"
        print(error_msg)
        return False, error_msg
        
    except socket.error as e:
        error_msg = f"Network error connecting to {ip}:{SSH_PORT}: {str(e)}"
        print(error_msg)
        return False, error_msg
        
    except Exception as e:
        error_msg = f"Unexpected error connecting to {ip}: {str(e)}"
        print(error_msg)
        return False, error_msg
        
    finally:
        # Always close the connection
        if ssh_client:
            try:
                ssh_client.close()
                print(f"SSH connection to {ip} closed")
            except:
                pass


def test_mikrotik_connection_detailed(ip):
    """
    Detailed connection test function
    """
    import socket
    
    print(f"=== Testing connection to {ip}:{SSH_PORT} ===")
    
    # Test 1: Basic network connectivity
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(10)
        result = sock.connect_ex((ip, SSH_PORT))
        sock.close()
        
        if result == 0:
            print("✓ Network connectivity: OK")
        else:
            print(f"✗ Network connectivity: FAILED (error code: {result})")
            return False, f"Cannot reach {ip}:{SSH_PORT}"
            
    except Exception as e:
        print(f"✗ Network connectivity: FAILED ({e})")
        return False, f"Network error: {e}"
    
    # Test 2: SSH connection and authentication
    try:
        ssh_client = paramiko.SSHClient()
        ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        
        ssh_client.connect(
            hostname=ip,
            username=MIKROTIK_USERNAME,
            password=MIKROTIK_PASSWORD,
            port=MIKROTIK_SSH_PORT,
            timeout=30,
            look_for_keys=False,
            allow_agent=False
        )
        
        print("✓ SSH authentication: OK")
        
        # Test 3: Execute a simple command
        stdin, stdout, stderr = ssh_client.exec_command('/system identity print', timeout=15)
        exit_status = stdout.channel.recv_exit_status()
        output = stdout.read().decode('utf-8', errors='ignore')
        
        ssh_client.close()
        
        if exit_status == 0:
            print("✓ Command execution: OK")
            print(f"Device identity: {output.strip()}")
            return True, "Connection test successful"
        else:
            error = stderr.read().decode('utf-8', errors='ignore')
            print(f"✗ Command execution: FAILED (exit status: {exit_status})")
            return False, f"Command failed: {error}"
            
    except paramiko.AuthenticationException:
        print("✗ SSH authentication: FAILED")
        return False, "Authentication failed - check username/password"
        
    except paramiko.SSHException as e:
        print(f"✗ SSH connection: FAILED ({e})")
        return False, f"SSH error: {e}"
        
    except Exception as e:
        print(f"✗ Connection test: FAILED ({e})")
        return False, f"Test failed: {e}"


def parse_mikrotik_firewall_output(output):
    """
    Parse MikroTik firewall address-list output into structured data
    Using the same logic as extract-suspend route
    """
    lines = output.strip().split('\n')
    extracted_data = []
    current_csid = None
    entry_number = 0
    
    for line in lines:
        # Extract CSID (SI-BK followed by 6 digits)
        csid_match = re.search(r'SI-BK\d{6}', line)
        if csid_match:
            current_csid = csid_match.group()
        
        # Extract IP address
        ip_match = re.search(r'\b(?:\d{1,3}\.){3}\d{1,3}\b', line)
        if ip_match:
            ip = ip_match.group()
            
            # Extract date if present in the line
            date_match = re.search(r'(\w{3}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2})', line)
            date_added = date_match.group(1) if date_match else 'N/A'
            
            # Create entry
            if current_csid:
                extracted_data.append({
                    "number": str(entry_number),
                    "csid": current_csid,
                    "ip": ip,
                    "date": date_added
                })
                current_csid = None  # Reset after pairing
            else:
                extracted_data.append({
                    "number": str(entry_number),
                    "csid": "N/A",
                    "ip": ip,
                    "date": date_added
                })
            
            entry_number += 1
    
    # Remove duplicates while preserving order
    seen = set()
    unique_data = []
    for item in extracted_data:
        key = (item["csid"], item["ip"])
        if key not in seen:
            seen.add(key)
            unique_data.append(item)
    
    # Sort by CSID (A to Z), putting "N/A" values at the end
    unique_data.sort(key=lambda x: (x["csid"] == "N/A", x["csid"]))
    
    return unique_data

def find_in_all_sheets(session_id, search_value):
    """Find search value in all sheets for a specific session"""
    if session_id not in workbook_storage:
        return "No data loaded"
    
    workbook_data = workbook_storage[session_id]['data']
    found_sheets = []
    search_str = str(search_value)
    
    for sheet_name, column_data in workbook_data.items():
        str_data = [str(val) for val in column_data if str(val) != 'nan']
        if search_str in str_data:
            found_sheets.append(sheet_name)
    
    if found_sheets:
        return ", ".join(found_sheets)
    else:
        return "Not Found"

@app.route('/manage-active.html')
def manage_active():
    """Route for the Manage Active - Batch CSID Status Checker page"""
    return render_template('manage-active.html')


@app.route('/')
def index():
    """Main page"""
    if 'session_id' not in session:
        session['session_id'] = str(uuid.uuid4())
    return render_template('home.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle Excel file upload"""
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if not file.filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'Please upload an Excel file (.xlsx or .xls)'}), 400
    
    try:
        # Read Excel file
        workbook_data = {}
        with pd.ExcelFile(file) as xls:
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                if len(df.columns) >= 4:  # Check if column D exists
                    column_d = df.iloc[:, 3].astype(str).dropna()  # Column D (index 3)
                    workbook_data[sheet_name] = column_d.tolist()
                else:
                    workbook_data[sheet_name] = []
        
        # Store in session storage
        session_id = session['session_id']
        workbook_storage[session_id] = {
            'data': workbook_data,
            'filename': secure_filename(file.filename),
            'upload_time': datetime.now().isoformat()
        }
        
        return jsonify({
            'success': True,
            'message': f'Loaded data from {len(workbook_data)} sheets',
            'sheets': list(workbook_data.keys())
        })
        
    except Exception as e:
        return jsonify({'error': f'Failed to load Excel file: {str(e)}'}), 500

@app.route('/add_csid', methods=['POST'])
def add_csid():
    """Add single CSID"""
    data = request.get_json()
    csid = data.get('csid', '').strip()
    
    if not csid:
        return jsonify({'error': 'CSID cannot be empty'}), 400
    
    session_id = session['session_id']
    found_sheets = find_in_all_sheets(session_id, csid)
    
    return jsonify({
        'success': True,
        'csid': csid,
        'found_sheets': found_sheets
    })

@app.route('/add_bulk_csids', methods=['POST'])
def add_bulk_csids():
    """Add multiple CSIDs"""
    data = request.get_json()
    csids_text = data.get('csids', '').strip()
    
    if not csids_text:
        return jsonify({'error': 'No CSIDs provided'}), 400
    
    # Parse CSIDs
    csids = re.split(r'[\s,;\n\t]+', csids_text)
    csids = [csid.strip() for csid in csids if csid.strip()]
    
    if not csids:
        return jsonify({'error': 'No valid CSIDs found'}), 400
    
    session_id = session['session_id']
    results = []
    
    for csid in csids:
        found_sheets = find_in_all_sheets(session_id, csid)
        results.append({
            'csid': csid,
            'found_sheets': found_sheets
        })
    
    return jsonify({
        'success': True,
        'results': results,
        'count': len(results)
    })

@app.route('/refresh_csids', methods=['POST'])
def refresh_csids():
    """Refresh existing CSIDs"""
    data = request.get_json()
    csids = data.get('csids', [])
    
    if not csids:
        return jsonify({'error': 'No CSIDs to refresh'}), 400
    
    session_id = session['session_id']
    results = []
    
    for csid in csids:
        found_sheets = find_in_all_sheets(session_id, csid)
        results.append({
            'csid': csid,
            'found_sheets': found_sheets
        })
    
    return jsonify({
        'success': True,
        'results': results
    })

@app.route('/export', methods=['POST'])
def export_data():
    """Export data to Excel"""
    data = request.get_json()
    export_data = data.get('data', [])
    
    if not export_data:
        return jsonify({'error': 'No data to export'}), 400
    
    try:
        # Create DataFrame
        df = pd.DataFrame(export_data)
        
        # Create Excel file in memory
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='CSID_Results')
        
        output.seek(0)
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'csid_results_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )
        
    except Exception as e:
        return jsonify({'error': f'Export failed: {str(e)}'}), 500

@app.route('/status')
def get_status():
    """Get current session status"""
    session_id = session.get('session_id')
    
    if session_id in workbook_storage:
        wb_info = workbook_storage[session_id]
        return jsonify({
            'loaded': True,
            'filename': wb_info['filename'],
            'sheets': list(wb_info['data'].keys()),
            'upload_time': wb_info['upload_time']
        })
    else:
        return jsonify({'loaded': False})

@app.route('/tracking-co.html')
def track_home():
    return render_template('tracking-co.html')

@app.route('/tracking-odp.html')
def track_odp():
    return render_template('tracking-odp.html')
    
@app.route('/tracking-ip.html')
def track_ip():
    return render_template('tracking-ip.html')

@app.route('/search-odp', methods=['POST'])
def search_odp():
    odp_id = request.form.get('odp_id')
    file = request.files['file']
    
    wb = openpyxl.load_workbook(file)
    results = []
    
    for sheet_name in wb.sheetnames:
        if sheet_name == "TRACK ODP":
            continue
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[1] and str(row[1]).strip() == odp_id.strip():
                results.append({
                    'ip': row[2],
                    'csid': row[3]
                })
    
    return jsonify(results)

@app.route('/search-ip', methods=['POST'])
def search_ip():
    csid_input = request.form.get('csid')
    file = request.files['file']
    
    # Parse multiple CSIDs - split by newlines, commas, or semicolons
    csids = []
    if csid_input:
        # Split by various delimiters and clean up
        csid_list = re.split(r'[,;\n\r]+', csid_input.strip())
        csids = [csid.strip() for csid in csid_list if csid.strip()]
    
    if not csids:
        return jsonify({'error': 'Please provide at least one CSID'})
    
    wb = openpyxl.load_workbook(file)
    
    # Build the lookup table first by scanning all sheets once
    ip_map = {}  # Dictionary: csid -> {'ip': ip_address, 'sheet': sheet_name}
    
    # Single pass through all sheets to build the lookup map
    for sheet_name in wb.sheetnames:
        if sheet_name == "TRACK ODP":
            continue
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[3]:  # Check if CSID column has value
                row_csid = str(row[3]).strip()
                if row_csid:  # Only add non-empty CSIDs
                    ip_map[row_csid] = {
                        'ip': row[2] if row[2] else 'N/A',
                        'sheet': sheet_name
                    }
    
    # Fast O(1) lookup for each searched CSID
    results = []
    found_csids = set()
    
    for search_csid in csids:
        if search_csid in ip_map:
            # Found the CSID
            data = ip_map[search_csid]
            results.append({
                'csid': search_csid,
                'ip': data['ip'],
                'sheet': data['sheet']
            })
            found_csids.add(search_csid)
    
    # Add entries for CSIDs that were not found
    not_found_csids = set(csids) - found_csids
    for csid in not_found_csids:
        results.append({
            'csid': csid,
            'ip': 'Not Found',
            'sheet': 'N/A'
        })
    
    # Sort results to show found ones first, then not found
    results.sort(key=lambda x: (x['ip'] == 'Not Found', x['csid']))
    
    return jsonify(results)

# MODIFIED: Changed route to serve automate-suspend.html instead of index.html
@app.route('/automate-suspend.html')
def automate_suspend():
    """Route for the SUSPEND CO automation page with MikroTik integration"""
    return render_template('automate-suspend.html')

@app.route('/extract-suspend', methods=['POST'])
def extract_suspend():
    """Extract CSID and IP addresses from SUSPEND CO text"""
    input_text = request.form.get('input_text', '')
    
    lines = input_text.strip().split('\n')
    extracted_data = []
    current_csid = None
    
    for line in lines:
        # Extract CSID (SI-BK followed by 6 digits)
        csid_match = re.search(r'SI-BK\d{6}', line)
        if csid_match:
            current_csid = csid_match.group()
        
        # Extract IP address
        ip_match = re.search(r'\b(?:\d{1,3}\.){3}\d{1,3}\b', line)
        if ip_match:
            ip = ip_match.group()
            if current_csid:
                extracted_data.append({"csid": current_csid, "ip": ip})
                current_csid = None  # Reset after pairing
            else:
                extracted_data.append({"csid": "N/A", "ip": ip})
    
    # Remove duplicates while preserving order
    seen = set()
    unique_data = []
    for item in extracted_data:
        key = (item["csid"], item["ip"])
        if key not in seen:
            seen.add(key)
            unique_data.append(item)
    
    # Sort by CSID (A to Z), putting "N/A" values at the end
    unique_data.sort(key=lambda x: (x["csid"] == "N/A", x["csid"]))
    
    return jsonify({"results": unique_data, "count": len(unique_data)})

@app.route('/get-sample-text', methods=['GET'])
def get_sample_text():
    """Return sample text for the SUSPEND CO automation"""
    sample_text = """0   ;;; Blocked: SI-BK003633
     SUSPEND_SOLNET                                      10.14.25.159                                                             jul/04/2025 10:02:53
 1   ;;; Blocked: SI-BK003831
     SUSPEND_SOLNET                                      10.14.25.230                                                             jul/05/2025 11:56:09"""
    return jsonify({"sample_text": sample_text})

@app.route('/reset_file', methods=['POST'])
def reset_file():
    """Reset the uploaded file data"""
    session_id = session.get('session_id')
    if session_id in workbook_storage:
        del workbook_storage[session_id]
    return jsonify({'success': True, 'message': 'File data reset successfully'})

# NEW MIKROTIK API ENDPOINTS
@app.route('/api/mikrotik/firewall-list', methods=['POST'])
def get_firewall_list():
    """
    API endpoint to get firewall address list from MikroTik device
    """
    try:
        data = request.get_json()
        
        if not data:
            return jsonify({'error': 'No data provided'}), 400
        
        ip = data.get('ip')
        command = data.get('command', '/ip firewall address-list print where list="SUSPEND_SOLNET"')
        
        if not ip:
            return jsonify({'error': 'IP address is required'}), 400
        
        # Connect to MikroTik and execute command
        success, output = connect_to_mikrotik(ip, command)
        
        if not success:
            return jsonify({'error': output}), 500
        
        # Parse the output
        results = parse_mikrotik_firewall_output(output)
        
        return jsonify({
            'success': True,
            'device_ip': ip,
            'results': results,
            'raw_output': output  # Include raw output for debugging
        })
        
    except Exception as e:
        return jsonify({'error': f'Server error: {str(e)}'}), 500

@app.route('/api/mikrotik/test-connection', methods=['POST'])
def test_connection():
    """
    Test connection to MikroTik device
    """
    try:
        data = request.get_json()
        ip = data.get('ip')
        
        if not ip:
            return jsonify({'error': 'IP address is required'}), 400
        
        # Test connection with a simple command
        success, output = connect_to_mikrotik(ip, '/system identity print')
        
        if success:
            return jsonify({
                'success': True,
                'message': 'Connection successful',
                'device_ip': ip,
                'identity': output.strip()
            })
        else:
            return jsonify({
                'success': False,
                'error': output
            }), 500
            
    except Exception as e:
        return jsonify({'error': f'Server error: {str(e)}'}), 500

@app.route('/api/mikrotik/list-names', methods=['POST'])
def get_list_names():
    """
    Get all available address-list names from MikroTik
    """
    try:
        data = request.get_json()
        ip = data.get('ip')
        
        if not ip:
            return jsonify({'error': 'IP address is required'}), 400
        
        # Get all address-lists to see available list names
        success, output = connect_to_mikrotik(ip, '/ip firewall address-list print')
        
        if not success:
            return jsonify({'error': output}), 500
        
        # Extract unique list names
        lines = output.split('\n')
        list_names = set()
        
        for line in lines:
            # Look for list= patterns
            list_match = re.search(r'list=([^\s]+)', line)
            if list_match:
                list_name = list_match.group(1).strip('"')
                list_names.add(list_name)
        
        return jsonify({
            'success': True,
            'device_ip': ip,
            'available_lists': sorted(list(list_names)),
            'raw_output': output
        })
        
    except Exception as e:
        return jsonify({'error': f'Server error: {str(e)}'}), 500

@app.route('/api/mikrotik/debug-output', methods=['POST'])
def debug_output():
    """
    Debug endpoint to see raw MikroTik output
    """
    try:
        data = request.get_json()
        ip = data.get('ip')
        command = data.get('command', '/ip firewall address-list print where list="SUSPEND_SOLNET"')
        
        if not ip:
            return jsonify({'error': 'IP address is required'}), 400
        
        # Connect to MikroTik and execute command
        success, output = connect_to_mikrotik(ip, command)
        
        if not success:
            return jsonify({'error': output}), 500
        
        # Return raw output split by lines for easier debugging
        lines = output.split('\n')
        numbered_lines = []
        for i, line in enumerate(lines):
            numbered_lines.append(f"{i:2d}: {repr(line)}")
        
        return jsonify({
            'success': True,
            'device_ip': ip,
            'raw_output': output,
            'lines': numbered_lines
        })
        
    except Exception as e:
        return jsonify({'error': f'Server error: {str(e)}'}), 500
        
@app.route('/api/mikrotik/batch-status-check', methods=['POST'])
def batch_status_check():
    try:
        data = request.get_json()
        
        # Debug logging
        app.logger.info(f"Received data: {data}")
        
        if not data or 'csids' not in data:
            return jsonify({'error': 'No CSIDs provided'}), 400
        
        csid_items = data['csids']  # Note: Using the same parameter name as frontend
        results = []
        
        # Group by device IP
        device_groups = {}
        for item in csid_items:
            ip = item.get('deviceIp')
            if ip and ip != 'Unknown':
                if ip not in device_groups:
                    device_groups[ip] = []
                device_groups[ip].append(item)
        
        # Process each device group
        for ip, items in device_groups.items():
            try:
                # First check suspended list (SUSPEND_SOLNET)
                suspend_success, suspend_output = connect_to_mikrotik(
                    ip,
                    '/ip firewall address-list print where list="SUSPEND_SOLNET"'
                )
                
                suspended_csids = []
                if suspend_success:
                    suspended_data = parse_mikrotik_firewall_output(suspend_output)
                    suspended_csids = [item['csid'] for item in suspended_data if item['csid'] != 'N/A']
                
                # Then check main address list with detailed info
                main_success, main_output = connect_to_mikrotik(
                    ip,
                    '/ip firewall address-list print detail'
                )
                
                address_list_csids = {}
                if main_success:
                    # Parse the detailed output to get enabled/disabled status
                    lines = main_output.strip().split('\n')
                    
                    current_csid = None
                    is_disabled = False
                    
                    for line in lines:
                        line_stripped = line.strip()
                        
                        # Skip empty lines
                        if not line_stripped:
                            continue
                        
                        # Look for CSID in current line
                        csid_match = re.search(r'SI-BK\d{6}', line_stripped)
                        if csid_match:
                            current_csid = csid_match.group()
                        
                        # Check if this is a status line (starts with number and optionally X)
                        # Format examples:
                        # "25 X ;;; SUDAH BONGKAR by Wawan"
                        # "26   ;;; Some comment"
                        # "27 X"
                        status_match = re.match(r'^\s*(\d+)\s*(X)?\s*(?:;;;.*)?$', line_stripped)
                        if status_match:
                            x_flag = status_match.group(2)
                            is_disabled = bool(x_flag)  # True if X is present, False otherwise
                        
                        # Check for explicit disabled status in any line
                        if 'disabled=yes' in line.lower():
                            is_disabled = True
                        elif 'disabled=no' in line.lower():
                            is_disabled = False
                        
                        # If we found a CSID and we're processing lines that belong to this entry
                        # Look for IP address to confirm this is a complete entry
                        ip_match = re.search(r'\b(?:\d{1,3}\.){3}\d{1,3}\b', line_stripped)
                        if current_csid and ip_match:
                            # Save the status for this CSID
                            address_list_csids[current_csid] = 'disabled' if is_disabled else 'enabled'
                            app.logger.info(f"Found CSID {current_csid} with status: {'disabled' if is_disabled else 'enabled'}")
                            # Reset for next entry
                            current_csid = None
                            is_disabled = False
                
                # Determine status for each CSID based on new logic
                for item in items:
                    csid = item['csid']
                    
                    # Step 1: Check if in SUSPEND_SOLNET
                    if csid in suspended_csids:
                        status = 'SUSPENDED'
                    # Step 2: Check if in address-list
                    elif csid in address_list_csids:
                        if address_list_csids[csid] == 'enabled':
                            status = 'ACTIVE'
                        else:
                            status = 'DISABLED'
                    # Step 3: Not found in either
                    else:
                        status = 'NOT_FOUND'
                    
                    app.logger.info(f"CSID {csid}: Final status = {status}")
                    
                    results.append({
                        'csid': csid,
                        'co': item['co'],
                        'deviceIp': ip,
                        'status': status,
                        'error': None
                    })
                    
            except Exception as e:
                app.logger.error(f"Error processing {ip}: {str(e)}")
                for item in items:
                    results.append({
                        'csid': item['csid'],
                        'co': item['co'],
                        'deviceIp': ip,
                        'status': 'NOT_FOUND',
                        'error': str(e)
                    })
        
        # Calculate statistics
        stats = {
            'active': len([r for r in results if r['status'] == 'ACTIVE']),
            'suspended': len([r for r in results if r['status'] == 'SUSPENDED']),
            'disabled': len([r for r in results if r['status'] == 'DISABLED']),
            'not_found': len([r for r in results if r['status'] == 'NOT_FOUND']),
            'total': len(results)
        }
        
        return jsonify({
            'success': True,
            'results': results,
            'stats': stats
        })
        
    except Exception as e:
        app.logger.error(f"Server error: {str(e)}")
        return jsonify({'error': f'Server error: {str(e)}'}), 500
        
# Optional: Add a route to serve devices configuration for the frontend
@app.route('/api/devices/co-mapping', methods=['GET'])
def get_co_mapping():
    """
    Get CO to IP mapping for the frontend
    """
    try:
        devices = load_devices_active()
        co_mapping = {}
        
        for device in devices:
            if device.get('co') and device.get('ip'):
                co_mapping[device['co']] = device['ip']
        
        return jsonify({
            'success': True,
            'co_mapping': co_mapping
        })
        
    except Exception as e:
        return jsonify({'error': f'Failed to load CO mapping: {str(e)}'}), 500

@app.route('/api/devices-suspend', methods=['GET'])
def get_devices_suspend():
    """
    Get list of available MikroTik devices from JSON file
    """
    devices = load_devices_suspend()    

    return jsonify({'devices': devices})

@app.route('/api/devices-active', methods=['GET'])
def get_devices_active():
    """
    Get list of available MikroTik devices from JSON file
    """
    devices = load_devices_active() 
    return jsonify({'devices': devices})

if __name__ == '__main__':
    print("Starting WEB TRACK CSID with MikroTik Integration...")
    print("Make sure to:")
    print("1. Install required packages: pip install flask paramiko flask-cors pandas openpyxl")
    print("2. Create devices.json file in the same directory as app.py")
    print("3. Ensure your MikroTik devices have SSH enabled")
    print("4. Update your automate-suspend.html template to include MikroTik functionality")
    
    show_startup_message()
    
    # Start browser in background
    if getattr(sys, 'frozen', False):
        threading.Thread(target=open_browser, daemon=True).start()
    
    app.run(debug=False, host='192.168.10.2', port=5050)