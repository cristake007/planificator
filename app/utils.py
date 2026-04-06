from datetime import datetime
from typing import List, Optional
import os
from config import Config

def validate_file_extension(filename: str) -> bool:
    """Check if the file extension is allowed."""
    return os.path.splitext(filename)[1].lower() in Config.ALLOWED_EXTENSIONS

def validate_year(year: int) -> bool:
    """Check if the year is within allowed range."""
    return Config.MIN_YEAR <= year <= Config.MAX_YEAR

def validate_date_format(date_str: str) -> bool:
    """Validate date string format (DD.MM.YYYY)."""
    try:
        datetime.strptime(date_str, '%d.%m.%Y')
        return True
    except ValueError:
        return False

def validate_holidays(holidays: List[str]) -> tuple[bool, Optional[str]]:
    """Validate list of holiday dates."""
    for date in holidays:
        if not date.strip():
            continue
        if not validate_date_format(date.strip()):
            return False, f"Invalid date format: {date}. Use DD.MM.YYYY"
    return True, None

def get_month_name(month_number: int) -> str:
    """Convert month number to name."""
    months = [
        'January', 'February', 'March', 'April',
        'May', 'June', 'July', 'August',
        'September', 'October', 'November', 'December'
    ]
    return months[month_number - 1]

def sanitize_filename(filename: str) -> str:
    """Sanitize filename for safe file operations."""
    # Remove any path components (on Unix and Windows)
    filename = os.path.basename(filename)

    # Remove any non-alphanumeric characters except periods and hyphens
    filename = ''.join(c for c in filename if c.isalnum() or c in '.-_')

    return filename

def validate_input_data(df) -> tuple[bool, Optional[str]]:
    """Validate the input dataframe structure and content."""
    required_cols = ['course_name', 'duration']

    # Check for required columns
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        return False, f"Missing required columns: {', '.join(missing_cols)}"

    # Check for empty dataframe
    if df.empty:
        return False, "Input file contains no data"

    # Validate duration values
    invalid_durations = df[~df['duration'].apply(lambda x: isinstance(x, (int, float)) and x > 0)]
    if not invalid_durations.empty:
        return False, "Duration must be a positive number"

    # Validate course names
    empty_names = df[df['course_name'].isna() | (df['course_name'] == '')]
    if not empty_names.empty:
        return False, "Course names cannot be empty"

    return True, None

def format_error_message(error: Exception) -> str:
    """Format error message for user display."""
    error_str = str(error)
    if 'Max retries exceeded' in error_str:
        return 'Server is temporarily unavailable. Please try again later.'
    if 'Invalid file format' in error_str:
        return 'Please upload a valid CSV or Excel file.'
    return error_str

def check_file_size(file_size: int) -> tuple[bool, Optional[str]]:
    """Check if file size is within allowed limit."""
    max_size = Config.MAX_CONTENT_LENGTH
    if file_size > max_size:
        max_size_mb = max_size / (1024 * 1024)
        return False, f"File size exceeds maximum limit of {max_size_mb}MB"
    return True, None