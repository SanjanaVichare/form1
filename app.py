from flask import Flask, render_template, request, redirect, url_for, jsonify
import os
import pandas as pd
from datetime import datetime, timedelta
import json

app = Flask(__name__)

EXCEL_FILE = 'bookings.xlsx'

# Initialize Excel file if it doesn't exist
def initialize_excel():
    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=[
            'Full Name', 'Email', 'Phone', 'Company', 'Setup', 'People', 'Experience',
            'Package', 'Addons', 'Date', 'Duration', 'Time Slot', 'Frequency',
            'Requirements', 'Referral', 'Submission Date'
        ])
        df.to_excel(EXCEL_FILE, index=False)

# Load booked slots from Excel with date filtering
def get_booked_slots(date=None):
    if not os.path.exists(EXCEL_FILE):
        initialize_excel()
        return []
    
    try:
        df = pd.read_excel(EXCEL_FILE)
        
        # Check if dataframe is empty
        if df.empty:
            return []
        
        # Filter by specific date if provided
        if date:
            df = df[df['Date'] == date]
        
        # Check if we have any bookings after filtering
        if df.empty:
            return []
        
        booked_slots = []
        for _, row in df.iterrows():
            try:
                time_slot = str(row['Time Slot']).split(' ')[0]  # Get just the time part (HH:MM)
                duration = int(row['Duration'])
                
                # Parse the time slot (format: "HH:MM")
                start_time = datetime.strptime(time_slot, '%H:%M')
                end_time = start_time + timedelta(hours=duration)
                
                booked_slots.append({
                    'start': time_slot,
                    'duration': duration,
                    'end': end_time.strftime('%H:%M')
                })
            except (ValueError, KeyError, TypeError) as e:
                print(f"Error processing row: {e}")
                continue
        
        return booked_slots
    
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []

# Check if a time slot is available
def is_time_slot_available(date, start_time, duration):
    booked_slots = get_booked_slots(date)
    
    # If no booked slots, it's available
    if not booked_slots:
        return True
    
    # Parse the requested time
    requested_start = datetime.strptime(start_time, '%H:%M')
    requested_end = requested_start + timedelta(hours=duration)
    
    for booked in booked_slots:
        booked_start = datetime.strptime(booked['start'], '%H:%M')
        booked_end = datetime.strptime(booked['end'], '%H:%M')
        
        # Check for overlap
        if (requested_start < booked_end and requested_end > booked_start):
            return False
    
    return True

# Save booking to Excel
def save_to_excel(data):
    # Ensure Excel file exists
    initialize_excel()
    
    df = pd.read_excel(EXCEL_FILE)
    
    # Handle multiple selections (packages and addons)
    packages = data.get('package', [])
    if isinstance(packages, list):
        packages = ', '.join(packages)
    
    addons = data.get('addons', [])
    if isinstance(addons, list):
        addons = ', '.join(addons)
    
    new_data = {
        'Full Name': data.get('name'),
        'Email': data.get('email'),
        'Phone': data.get('phone'),
        'Company': data.get('company'),
        'Setup': data.get('setup'),
        'People': data.get('people'),
        'Experience': data.get('experience'),
        'Package': packages,
        'Addons': addons,
        'Date': data.get('date'),
        'Duration': int(data.get('duration', 1)),
        'Time Slot': data.get('time_slot'),
        'Frequency': data.get('frequency'),
        'Requirements': data.get('requirements'),
        'Referral': data.get('referral'),
        'Submission Date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }
    
    new_row = pd.DataFrame([new_data])
    df = pd.concat([df, new_row], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)

@app.route('/')
def index():
    return render_template('form.html')

@app.route('/api/booked_slots')
def get_booked_slots_api():
    date = request.args.get('date')
    booked_slots = get_booked_slots(date)
    print(f"Returning booked slots for {date}: {booked_slots}")  # Debug log
    return jsonify(booked_slots)

@app.route('/api/check_availability')
def check_availability():
    date = request.args.get('date')
    start_time = request.args.get('start_time')
    duration = request.args.get('duration', type=int)
    
    if not all([date, start_time, duration]):
        return jsonify({'available': False, 'error': 'Missing parameters'})
    
    available = is_time_slot_available(date, start_time, duration)
    return jsonify({'available': available})

@app.route('/submit', methods=['POST'])
def submit():
    try:
        # Get form data
        data = {
            'name': request.form.get('name'),
            'email': request.form.get('email'),
            'phone': request.form.get('phone'),
            'company': request.form.get('company'),
            'setup': request.form.get('setup'),
            'people': request.form.get('people'),
            'experience': request.form.get('experience'),
            'package': request.form.getlist('package'),  # Multiple selection
            'addons': request.form.getlist('addons'),    # Multiple selection
            'date': request.form.get('date'),
            'duration': request.form.get('duration'),
            'time_slot': request.form.get('time_slot'),
            'frequency': request.form.get('frequency'),
            'requirements': request.form.get('requirements'),
            'referral': request.form.get('referral')
        }
        
        # Validate required fields
        required_fields = ['name', 'email', 'phone', 'setup', 'people', 'date', 'duration', 'time_slot']
        for field in required_fields:
            if not data.get(field):
                return jsonify({'success': False, 'error': f'Missing required field: {field}'}), 400
        
        # Extract just the time part from time_slot (in case it includes duration)
        time_part = data['time_slot'].split(' ')[0]
        
        # Validate time slot availability
        if not is_time_slot_available(data['date'], time_part, int(data['duration'])):
            return jsonify({'success': False, 'error': 'Selected time slot is no longer available'}), 400
        
        # Save to Excel
        save_to_excel(data)
        
        # Send email notification (you can integrate this later)
        # send_email_notification(data)
        
        return jsonify({'success': True, 'message': 'Booking submitted successfully!'})
    
    except Exception as e:
        print(f"Error in submit: {e}")  # Debug log
        return jsonify({'success': False, 'error': str(e)}), 500

# Email sending function (placeholder)
def send_email_notification(data):
    # This is where you would integrate your email sending logic
    # You can use Flask-Mail, smtplib, or a service like SendGrid
    print(f"Email would be sent to inwmhstudios@gmail.com with data: {data}")
    pass

if __name__ == '__main__':
    initialize_excel()
    app.run(debug=True)
