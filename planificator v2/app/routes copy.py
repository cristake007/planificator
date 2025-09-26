from flask import Blueprint, render_template, request, jsonify, send_file
import pandas as pd
import random
import os
import io
from datetime import datetime, timedelta
from typing import List, Dict, Any
import xml.etree.ElementTree as ET
from xml.dom import minidom

from app.scheduler import CourseScheduler
from app.file_handlers import read_input_file, create_excel_export

scheduler_bp = Blueprint('scheduler', __name__)

def create_xml_export(schedule: List[Dict[str, Any]], year: int) -> str:
    """Create XML file from schedule data."""
    root = ET.Element("events")
    
    for idx, event in enumerate(schedule, start=1):
        # Parse the date range
        if '-' in event['date_range']:
            start_day, end_full = event['date_range'].split('-')
            start_date = f"{year}-{str(event['month']).zfill(2)}-{start_day.zfill(2)}"
            end_date = end_full.split('.')
            end_date = f"{year}-{end_date[1].zfill(2)}-{end_date[0].zfill(2)}"
        else:
            # Single day event
            date_parts = event['date_range'].split('.')
            start_date = end_date = f"{year}-{date_parts[1].zfill(2)}-{date_parts[0].zfill(2)}"

        item = ET.SubElement(root, "item")
        
        # ID and basic info
        ET.SubElement(item, "ID").text = str(20000 + idx)
        ET.SubElement(item, "title").text = event['course_name']
        ET.SubElement(item, "content").text = ""
        
        # Post details
        post = ET.SubElement(item, "post")
        ET.SubElement(post, "ID").text = str(20000 + idx)
        ET.SubElement(post, "post_author").text = "5"
        
        # Calculate timestamps
        start_datetime = datetime.strptime(start_date, "%Y-%m-%d")
        end_datetime = datetime.strptime(end_date, "%Y-%m-%d")
        
        post_date = start_datetime.strftime("%Y-%m-%d %H:%M:%S")
        ET.SubElement(post, "post_date").text = post_date
        ET.SubElement(post, "post_date_gmt").text = (start_datetime - timedelta(hours=3)).strftime("%Y-%m-%d %H:%M:%S")
        ET.SubElement(post, "post_title").text = event['course_name']
        ET.SubElement(post, "post_status").text = "draft"
        
        # Meta information
        meta = ET.SubElement(item, "meta")
        date = ET.SubElement(meta, "mec_date")
        
        
        # Additional meta information
        ET.SubElement(meta, "mec_color").text = ""
        ET.SubElement(meta, "mec_location_id").text = "1"
        ET.SubElement(meta, "mec_dont_show_map").text = "0"
        ET.SubElement(meta, "mec_organizer_id").text = "1"
        ET.SubElement(meta, "mec_read_more").text = event.get('link', 'https://www.google.com')

        # Extract link, default to empty string if not provided
        link = event.get('link', 'https://www.google.com')
        
        # Additional meta information
        ET.SubElement(meta, "mec_color").text = ""
        ET.SubElement(meta, "mec_location_id").text = "1"
        ET.SubElement(meta, "mec_dont_show_map").text = "0"
        ET.SubElement(meta, "mec_organizer_id").text = "1"
        ET.SubElement(meta, "mec_read_more").text = link
        
        # Additional required fields
        ET.SubElement(meta, "mec_allday").text = "1"
        ET.SubElement(meta, "event_past").text = "1"
        
        # Time information
        time = ET.SubElement(item, "time")
        ET.SubElement(time, "start").text = "All Day"
        ET.SubElement(time, "end").text = ""
        ET.SubElement(time, "start_raw").text = "8:00 am"
        ET.SubElement(time, "end_raw").text = "6:00 pm"
        
        # Read more link (multiple locations to ensure visibility)
        ET.SubElement(item, "read_more_link").text = link
        mec_read_more = ET.SubElement(item, "mec_read_more")
        mec_read_more.text = link
        
        # Start date info
        start = ET.SubElement(date, "start")
        ET.SubElement(start, "date").text = start_date
        ET.SubElement(start, "hour").text = "8"
        ET.SubElement(start, "minutes").text = "0"
        ET.SubElement(start, "ampm").text = "AM"
        
        # End date info
        end = ET.SubElement(date, "end")
        ET.SubElement(end, "date").text = end_date
        ET.SubElement(end, "hour").text = "6"
        ET.SubElement(end, "minutes").text = "0"
        ET.SubElement(end, "ampm").text = "PM"
        
        ET.SubElement(meta, "mec_start_date").text = start_date
        ET.SubElement(meta, "mec_start_time_hour").text = "8"
        ET.SubElement(meta, "mec_start_time_minutes").text = "00"
        ET.SubElement(meta, "mec_start_time_ampm").text = "AM"
        ET.SubElement(meta, "mec_end_date").text = end_date
        ET.SubElement(meta, "mec_end_time_hour").text = "6"
        ET.SubElement(meta, "mec_end_time_minutes").text = "00"
        ET.SubElement(meta, "mec_end_time_ampm").text = "PM"
        
        # Additional required fields
        ET.SubElement(meta, "mec_allday").text = "1"
        ET.SubElement(meta, "event_past").text = "1"
        
        # Time information
        time = ET.SubElement(item, "time")
        ET.SubElement(time, "start").text = "All Day"
        ET.SubElement(time, "end").text = ""
        ET.SubElement(time, "start_raw").text = "8:00 am"
        ET.SubElement(time, "end_raw").text = "6:00 pm"
        
        # # Categories
        # categories = ET.SubElement(item, "categories")
        # cat_item = ET.SubElement(categories, "item")
        # ET.SubElement(cat_item, "id").text = "76"
        # ET.SubElement(cat_item, "name").text = "Management si soft skills"

    # Convert to string with pretty printing
    xmlstr = minidom.parseString(ET.tostring(root)).toprettyxml(indent="    ")
    return xmlstr



@scheduler_bp.route('/')
def index():
    """Render the main page."""
    return render_template('dashboard.html')

@scheduler_bp.route('/generator-perioade')
def generator_perioade():
    """Render the schedule generator page."""
    return render_template('generator_perioade.html')  # Separate generator page




@scheduler_bp.route('/generate_schedule', methods=['POST'])
def generate_schedule():
    """Generate schedule for selected months."""
    try:
        # Get input file and year
        file = request.files['input_file']
        year = int(request.form['year'])
        randomness_level = int(request.form.get('randomness', 5))

        # Get selected months
        selected_months = request.form.get('months', '')
        if not selected_months:
            return jsonify({'success': False, 'error': 'Please select at least one month'})
        
        selected_months = [int(m) for m in selected_months.split(',')]
        if not selected_months:
            return jsonify({'success': False, 'error': 'Please select at least one month'})

        # Get holidays from the form
        holidays = request.form.get('holidays', '').split(',')
        holidays = [h.strip() for h in holidays if h.strip()]

        # Read file
        file_extension = os.path.splitext(file.filename)[1]
        df = read_input_file(file.read(), file_extension)

        if df.empty:
            return jsonify({'success': False, 'error': 'Input file is empty or missing required data.'})

        # Ensure columns are present
        required_columns = ['Title', 'Permalink', 'Durata Curs']
        for col in required_columns:
            if col not in df.columns:
                return jsonify({'success': False, 'error': f'Missing required column: {col}'})

        # Initialize scheduler
        scheduler = CourseScheduler(year, holidays)

        schedule = []

        # Only process selected months
        for month in selected_months:
            scheduled_dates = set()

            # Process courses in original order
            for _, row in df.iterrows():
                title = row['Title']
                permalink = row['Permalink']
                durata_curs_raw = row['Durata Curs']

                # Extract integer duration from 'Durata Curs' (e.g., "3 zile" -> 3)
                match = re.search(r'(\d+)', durata_curs_raw)
                durata_curs = int(match.group(1)) if match else 0

                available_dates = scheduler.get_available_start_days(month, durata_curs)

                if available_dates:
                    # Remove dates too close to already scheduled dates based on randomness
                    min_gap = max(1, (11 - randomness_level))
                    filtered_dates = [
                        date for date in available_dates
                        if not any(abs((date - scheduled).days) < min_gap for scheduled in scheduled_dates)
                    ]

                    # If no dates available with minimum gap, use any available date
                    dates_to_use = filtered_dates if filtered_dates else available_dates

                    if dates_to_use:
                        if randomness_level > 7:
                            start_date = random.choice(dates_to_use)
                        else:
                            num_dates = len(dates_to_use)
                            mid_point = num_dates // 2
                            weights = [1.0 - (abs(i - mid_point) / num_dates) * 0.5 for i in range(num_dates)]
                            start_date = random.choices(dates_to_use, weights=weights, k=1)[0]

                        schedule.append({
                            'Title': title,
                            'Permalink': permalink,
                            'Durata Curs': durata_curs,
                            'date_range': scheduler.format_date_range(start_date, durata_curs),
                            'month': month,
                        })

                        scheduled_dates.add(start_date)

        # Sort by month, then by course name within each month
        schedule.sort(key=lambda x: (x['month'], x['Title']))

        # Debug: Output course summary
        total_unique_courses = len(set(item['Title'] for item in schedule))
        total_scheduled_sessions = len(schedule)
        print(f"Total unique courses: {total_unique_courses}, Total scheduled sessions: {total_scheduled_sessions}")

        return jsonify({'success': True, 'schedule': schedule})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@scheduler_bp.route('/export_schedule', methods=['POST'])
def export_schedule():
    """Export schedule to Excel file."""
    try:
        schedule_data = request.json.get('schedule', [])
        year = int(request.json.get('year', datetime.now().year))
        holidays = request.json.get('holidays', [])

        excel_data = create_excel_export(schedule_data, year, holidays)

        return send_file(
            io.BytesIO(excel_data),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'course_schedule_{year}.xlsx'
        )
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})
    
@scheduler_bp.route('/xml-formatter')
def xml_formatter():
    """Render the XML formatter page."""
    return render_template('xml_formatter.html')

@scheduler_bp.route('/format-xml', methods=['POST'])
def format_xml():
    """Generate XML from Excel input."""
    try:
        print("format-xml endpoint hit")  # Debug print
        file = request.files.get('input_file')
        if not file:
            print("No file in request")  # Debug print
            return jsonify({'error': 'No file provided'}), 400

        print(f"File received: {file.filename}")  # Debug print

        year = datetime.now().year  # Get current year
        df = pd.read_excel(io.BytesIO(file.read()))

        schedule = []
        
        # Exclude non-month columns and preserve course details
        month_columns = [col for col in df.columns if col not in ['Course Name', 'Link', 'Course Link']]

        for _, row in df.iterrows():
            for month_col in month_columns:
                try:
                    month_number = datetime.strptime(month_col, '%B').month  # Parse month name safely
                except ValueError:
                    continue  # Skip columns that are not valid month names

                if pd.notna(row[month_col]):
                    date_range = str(row[month_col]).strip()
                    if date_range:
                        schedule.append({
                            'course_name': row['Course Name'],
                            'date_range': date_range,
                            'month': month_number,
                            'link': row.get('Link', row.get('Course Link', ''))  # Use course link if available
                        })

        xml_data = create_xml_export(schedule, year)

        return send_file(
            io.BytesIO(xml_data.encode('utf-8')),
            mimetype='application/xml',
            as_attachment=True,
            download_name=f'formatted_courses_{year}.xml'
        )

    except Exception as e:
        print(f"Error in format_xml: {str(e)}")  # Debug print
        return jsonify({'error': str(e)}), 400
