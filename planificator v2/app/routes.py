from flask import Blueprint, render_template, request, jsonify, send_file
import pandas as pd
import random
import os
import io
import calendar
from datetime import datetime, timedelta
import re
from typing import List, Dict, Any
import xml.etree.ElementTree as ET
from xml.dom import minidom
from app.file_handlers import convert_word_to_excel, process_word_file, read_input_file, create_excel_export
from rapidfuzz import process, fuzz
from app.scheduler import CourseScheduler
scheduler_bp = Blueprint('scheduler', __name__)

import xml.etree.ElementTree as ET
from xml.dom import minidom
from typing import List, Dict

def create_xml_export(schedule: List[Dict[str, Any]], year: int) -> str:
    """Create XML file from schedule data with course-specific period numbering."""
    try:
        def parse_date_range(date_range: str) -> tuple[str, str]:
            """Parse date ranges like 05.02.2026 or 05-06.02.2026 into ISO start/end dates."""
            normalized = (date_range or "").strip()

            # Single day format: dd.mm.yyyy
            single_day = re.match(r'^(\d{1,2})\.(\d{1,2})\.(\d{4})$', normalized)
            if single_day:
                day, month, parsed_year = single_day.groups()
                iso_date = f"{parsed_year}-{month.zfill(2)}-{day.zfill(2)}"
                return iso_date, iso_date

            # Multi-day same month format: dd-dd.mm.yyyy
            multi_day_same_month = re.match(r'^(\d{1,2})\s*-\s*(\d{1,2})\.(\d{1,2})\.(\d{4})$', normalized)
            if multi_day_same_month:
                start_day, end_day, month, parsed_year = multi_day_same_month.groups()
                start_iso = f"{parsed_year}-{month.zfill(2)}-{start_day.zfill(2)}"
                end_iso = f"{parsed_year}-{month.zfill(2)}-{end_day.zfill(2)}"
                return start_iso, end_iso

            # Multi-day explicit month format: dd.mm-dd.mm.yyyy
            multi_day_with_month = re.match(r'^(\d{1,2})\.(\d{1,2})\s*-\s*(\d{1,2})\.(\d{1,2})\.(\d{4})$', normalized)
            if multi_day_with_month:
                start_day, start_month, end_day, end_month, parsed_year = multi_day_with_month.groups()
                start_iso = f"{parsed_year}-{start_month.zfill(2)}-{start_day.zfill(2)}"
                end_iso = f"{parsed_year}-{end_month.zfill(2)}-{end_day.zfill(2)}"
                return start_iso, end_iso

            raise ValueError(f"Unsupported date format: {date_range}")

        root = ET.Element("events")
        
        # Group events by course name to track period numbers
        courses = {}
        for event in schedule:
            course_name = event['course_name']
            if course_name not in courses:
                courses[course_name] = []
            courses[course_name].append(event)
        
        event_id = 20000  # Starting ID for events
        
        # Process each course and its periods
        for course_name, course_events in courses.items():
            # Sort events by date if needed
            for period_idx, event in enumerate(course_events, start=1):
                try:
                    event_id += 1  # Increment ID for each event
                    date_str = event['date_range'].strip()
                    start_date, end_date = parse_date_range(date_str)

                    # MEC-compatible datetime meta values
                    start_day_seconds = 8 * 3600
                    end_day_seconds = 18 * 3600
                    start_datetime = f"{start_date} 08:00 AM"
                    end_datetime = f"{end_date} 06:00 PM"

                    # Create XML item
                    item = ET.SubElement(root, "item")
                    
                    # ID and title
                    ET.SubElement(item, "ID").text = str(event_id)
                    ET.SubElement(item, "title").text = course_name
                    ET.SubElement(item, "content").text = ""

                    # Post details
                    post = ET.SubElement(item, "post")
                    ET.SubElement(post, "ID").text = str(event_id)
                    ET.SubElement(post, "post_author").text = "5"
                    ET.SubElement(post, "post_date").text = f"{start_date} 00:00:00"
                    ET.SubElement(post, "post_date_gmt").text = f"{start_date} 00:00:00"
                    ET.SubElement(post, "post_title").text = course_name
                    ET.SubElement(post, "post_status").text = "draft"

                    # Meta information with course-specific period numbering
                    meta = ET.SubElement(item, "meta")
                    ET.SubElement(meta, "mec_more_info_title").text = f"perioada {period_idx}"
                    ET.SubElement(meta, "mec_read_more").text = event.get('permalink', '')
                    ET.SubElement(meta, "mec_color").text = ""
                    ET.SubElement(meta, "mec_location_id").text = "1"
                    ET.SubElement(meta, "mec_organizer_id").text = "1"
                    ET.SubElement(meta, "mec_allday").text = "1"  # Set as all-day event
                    ET.SubElement(meta, "mec_start_date").text = start_date
                    ET.SubElement(meta, "mec_start_time_hour").text = "8"
                    ET.SubElement(meta, "mec_start_time_minutes").text = "00"
                    ET.SubElement(meta, "mec_start_time_ampm").text = "AM"
                    ET.SubElement(meta, "mec_start_day_seconds").text = str(start_day_seconds)
                    ET.SubElement(meta, "mec_start_datetime").text = start_datetime
                    ET.SubElement(meta, "mec_end_date").text = end_date
                    ET.SubElement(meta, "mec_end_time_hour").text = "6"
                    ET.SubElement(meta, "mec_end_time_minutes").text = "00"
                    ET.SubElement(meta, "mec_end_time_ampm").text = "PM"
                    ET.SubElement(meta, "mec_end_day_seconds").text = str(end_day_seconds)
                    ET.SubElement(meta, "mec_end_datetime").text = end_datetime
                    ET.SubElement(meta, "mec_repeat_status").text = "0"
                    ET.SubElement(meta, "event_past").text = "1"

                    # mec_date with correct structure
                    mec_date = ET.SubElement(meta, "mec_date")
                    
                    start = ET.SubElement(mec_date, "start")
                    ET.SubElement(start, "date").text = start_date
                    ET.SubElement(start, "hour").text = "8"
                    ET.SubElement(start, "minutes").text = "00"
                    ET.SubElement(start, "ampm").text = "AM"

                    end = ET.SubElement(mec_date, "end")
                    ET.SubElement(end, "date").text = end_date
                    ET.SubElement(end, "hour").text = "6"
                    ET.SubElement(end, "minutes").text = "00"
                    ET.SubElement(end, "ampm").text = "PM"

                    # Set all-day inside mec_date
                    ET.SubElement(mec_date, "allday").text = "1"

                    # Explicit MEC/time blocks used by shortcode rendering in some imports.
                    mec_block = ET.SubElement(item, "mec")
                    ET.SubElement(mec_block, "id").text = ""
                    ET.SubElement(mec_block, "post_id").text = str(event_id)
                    ET.SubElement(mec_block, "start").text = start_date
                    ET.SubElement(mec_block, "end").text = end_date
                    ET.SubElement(mec_block, "repeat").text = "0"
                    ET.SubElement(mec_block, "time_start").text = str(start_day_seconds)
                    ET.SubElement(mec_block, "time_end").text = str(end_day_seconds)

                    time = ET.SubElement(item, "time")
                    ET.SubElement(time, "start").text = "All Day"
                    ET.SubElement(time, "end").text = ""
                    ET.SubElement(time, "start_raw").text = "8:00 am"
                    ET.SubElement(time, "end_raw").text = "6:00 pm"
                    ET.SubElement(time, "start_timestamp").text = str(int(datetime.strptime(start_date, "%Y-%m-%d").timestamp()) + start_day_seconds)
                    ET.SubElement(time, "end_timestamp").text = str(int(datetime.strptime(end_date, "%Y-%m-%d").timestamp()) + end_day_seconds)

                except Exception as e:
                    print(f"Error processing event for {course_name}, period {period_idx}: {str(e)}")
                    continue

        # Convert to XML string with pretty-printing
        xml_str = minidom.parseString(ET.tostring(root)).toprettyxml(indent="    ")
        return xml_str
        
    except Exception as e:
        print(f"Error in create_xml_export: {str(e)}")
        raise






@scheduler_bp.route('/')
def index():
    """Render the main page."""
    return render_template('dashboard.html')

@scheduler_bp.route('/generator-perioade')
def generator_perioade():
    """Render the schedule generator page."""
    return render_template('generator_perioade.html')  # Separate generator page

@scheduler_bp.route('/word-to-excel')
def word_to_excel():
    """Render the Word to Excel Converter page."""
    return render_template('word_converter.html')

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

        # Initialize scheduler
        scheduler = CourseScheduler(year, holidays)

        schedule = []

        # Only process selected months
        unscheduled_courses = {}
        for month in selected_months:
            scheduled_dates = set()
            month_unscheduled = set()

            # Process courses in original order
            for _, row in df.iterrows():
                duration = int(row['duration'])
                available_dates = scheduler.get_available_start_days(month, duration)

                if available_dates:
                    # Remove dates too close to already scheduled dates based on randomness
                    min_gap = max(1, (11 - randomness_level))
                    filtered_dates = []

                    for date in available_dates:
                        if not any(abs((date - scheduled).days) < min_gap
                                 for scheduled in scheduled_dates):
                            filtered_dates.append(date)

                    # If no dates available with minimum gap, use any available date
                    dates_to_use = filtered_dates if filtered_dates else available_dates

                    if dates_to_use:
                        if randomness_level > 7:
                            start_date = random.choice(dates_to_use)
                        else:
                            # Calculate weights based on position and randomness
                            num_dates = len(dates_to_use)
                            weights = []

                            for i in range(num_dates):
                                if randomness_level <= 3:
                                    if i < 5:  # First week
                                        weight = 0.5
                                    else:
                                        weight = 1.0
                                elif randomness_level <= 6:
                                    mid_point = num_dates // 2
                                    weight = 1.0 - (abs(i - mid_point) / num_dates) * 0.5
                                else:
                                    weight = 0.8 + random.random() * 0.4

                                weights.append(weight)

                            start_date = random.choices(dates_to_use, weights=weights, k=1)[0]

                        schedule.append({
                            'Title': row['Title'],
                            'Permalink': row['Permalink'],
                            'Durata Curs': row['Durata Curs'],
                            'date_range': scheduler.format_date_range(start_date, duration),
                            'month': month,
                            'original_order': int(row['original_order']),
                        })

                        scheduled_dates.add(start_date)
                    else:
                        course_identifier = str(row['Title']).strip() or f"Course #{int(row['original_order'])}"
                        month_unscheduled.add(course_identifier)
                else:
                    course_identifier = str(row['Title']).strip() or f"Course #{int(row['original_order'])}"
                    month_unscheduled.add(course_identifier)

            if month_unscheduled:
                unscheduled_courses[month] = sorted(month_unscheduled)

        if unscheduled_courses:
            month_messages = []
            for month, courses in unscheduled_courses.items():
                month_name = calendar.month_name[month]
                month_messages.append(f"{month_name}: {', '.join(courses)}")

            error_message = (
                "Unable to schedule all courses with the current constraints. "
                "The following courses had no available dates:\n" + "\n".join(month_messages)
            )
            return jsonify({
                'success': False,
                'error': error_message,
                'unscheduled_courses': unscheduled_courses
            }), 400

        # Sort by month, then by course name within each month
        schedule.sort(key=lambda x: x.get('original_order', 0))

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
    """Generate XML from CSV/Excel input with improved date parsing."""
    try:
        print("format-xml endpoint hit")
        file = request.files.get('input_file')
        if not file:
            print("No file in request")
            return jsonify({'error': 'No file provided'}), 400

        print(f"File received: {file.filename}")
        year = datetime.now().year

        file_extension = os.path.splitext(file.filename or '')[1].lower()
        supported_extensions = {'.xlsx', '.xls', '.csv'}
        if file_extension not in supported_extensions:
            return jsonify({
                'error': 'Unsupported file format. Please upload .xlsx, .xls, or .csv.'
            }), 400

        file_bytes = io.BytesIO(file.read())

        # Read input file
        try:
            if file_extension == '.csv':
                df = pd.read_csv(file_bytes, dtype=str, encoding='utf-8-sig', sep=None, engine='python')
            else:
                df = pd.read_excel(file_bytes, dtype=str)
            print("Input columns:", df.columns.tolist())
        except Exception as input_error:
            print(f"Input reading error: {str(input_error)}")
            return jsonify({'error': f'Error reading input file: {str(input_error)}'}), 400

        normalized_columns = {
            str(col).strip().lower(): col
            for col in df.columns
        }

        # Support both direct exports from the generator and legacy files.
        required_columns = {
            'title': 'Title',
            'permalink': 'Permalink',
        }
        missing_columns = [
            canonical_name
            for source_name, canonical_name in required_columns.items()
            if source_name not in normalized_columns
        ]
        if missing_columns:
            return jsonify({'error': f'Missing required columns: {", ".join(missing_columns)}'}), 400

        rename_map = {
            normalized_columns[source_name]: canonical_name
            for source_name, canonical_name in required_columns.items()
        }
        df = df.rename(columns=rename_map)

        month_aliases = {
            'january': 'January',
            'february': 'February',
            'march': 'March',
            'april': 'April',
            'may': 'May',
            'june': 'June',
            'july': 'July',
            'august': 'August',
            'september': 'September',
            'october': 'October',
            'november': 'November',
            'december': 'December',
            'luna 1': 'Luna 1',
            'luna 2': 'Luna 2',
            'luna 3': 'Luna 3',
            'luna 4': 'Luna 4',
            'luna 5': 'Luna 5',
            'luna 6': 'Luna 6',
            'luna 7': 'Luna 7',
            'luna 8': 'Luna 8',
            'luna 9': 'Luna 9',
            'luna 10': 'Luna 10',
            'luna 11': 'Luna 11',
            'luna 12': 'Luna 12',
        }

        month_columns = [
            original_name
            for normalized_name, original_name in normalized_columns.items()
            if normalized_name in month_aliases
        ]
        if not month_columns:
            return jsonify({
                'error': 'No supported date columns found. Use month columns (January-December) or Luna columns (Luna 1-Luna 12).'
            }), 400

        print(f"Found month columns: {month_columns}")
        
        # Process each row and create schedule data
        schedule = []
        for idx, row in df.iterrows():
            try:
                title = str(row['Title']).strip()
                permalink = str(row['Permalink']).strip()
                
                if not title:
                    print(f"Skipping row {idx}: missing title")
                    continue
                    
                for month_col in month_columns:
                    date_value = str(row[month_col]).strip()
                    if pd.notna(date_value) and date_value and date_value.lower() != 'nan':
                        print(f"Processing {title} for {month_col}: {date_value}")
                        
                        # Create a complete event entry
                        schedule.append({
                            'course_name': title,
                            'date_range': date_value,
                            'permalink': permalink
                        })
                    else:
                        print(f"No date for {title} in {month_col}")
            except Exception as row_error:
                print(f"Error processing row {idx}: {str(row_error)}")
                continue

        if not schedule:
            return jsonify({'error': 'No valid course data found in the input file'}), 400

        print(f"Successfully created {len(schedule)} events")
        
        # Generate XML with the fixed create_xml_export function
        xml_data = create_xml_export(schedule, year)

        return send_file(
            io.BytesIO(xml_data.encode('utf-8')),
            mimetype='application/xml',
            as_attachment=True,
            download_name=f'formatted_courses_{year}.xml'
        )

    except Exception as e:
        print(f"Error in format_xml: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 400

from rapidfuzz import process, fuzz

@scheduler_bp.route('/convert_word', methods=['POST'])
def convert_word():
    """Match Excel/CSV courses with Word document date ranges."""
    try:
        word_file = request.files.get('word_file')
        permalinks_file = request.files.get('permalinks_file')

        if not word_file or not word_file.filename.endswith('.docx'):
            return jsonify({'error': 'Invalid Word file. Please upload a .docx file'}), 400

        # Updated file extension check to include CSV
        if not permalinks_file or not permalinks_file.filename.lower().endswith(('.xlsx', '.xls', '.csv')):
            return jsonify({'error': 'Invalid file format. Please upload an .xlsx, .xls, or .csv file'}), 400

        # First read the Excel/CSV file as our base
        file_extension = permalinks_file.filename.lower().split('.')[-1]
        if file_extension == 'csv':
            base_df = pd.read_csv(permalinks_file)
        else:
            base_df = pd.read_excel(permalinks_file)

        # Validate base file structure
        if 'Title' not in base_df.columns or 'Permalink' not in base_df.columns:
            return jsonify({'error': 'Input file must contain "Title" and "Permalink" columns'}), 400

        # Get date ranges from Word file
        word_dates = process_word_file(word_file.read())

        # Add Luna columns to base DataFrame
        base_df['Luna 1'] = ''
        base_df['Luna 2'] = ''
        base_df['Luna 3'] = ''

        # Match and fill in Luna values using fuzzy matching
        for excel_idx, excel_row in base_df.iterrows():
            excel_title = excel_row['Title']
            best_match = process.extractOne(
                excel_title,
                list(word_dates.keys()),
                scorer=fuzz.ratio,
                score_cutoff=95
            )
            
            if best_match:
                word_title = best_match[0]
                luna_values = word_dates[word_title]
                for luna_col in ['Luna 1', 'Luna 2', 'Luna 3']:
                    base_df.at[excel_idx, luna_col] = luna_values[luna_col]

        # Convert to Excel
        output_excel = convert_to_excel(base_df)
        return send_file(
            io.BytesIO(output_excel),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='matched_courses.xlsx'
        )

    except Exception as e:
        import traceback
        print(traceback.format_exc())  # For debugging
        return jsonify({'error': str(e)}), 400


def convert_to_excel(df: pd.DataFrame) -> bytes:
    """Helper function to convert DataFrame to Excel format."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Courses')
    output.seek(0)
    return output.getvalue()
