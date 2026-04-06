class Config:
    # Flask configuration
    SECRET_KEY = 'your-secret-key'  # Change this in production
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB max file size

    # Application specific settings
    ALLOWED_EXTENSIONS = {'.csv', '.xlsx', '.xls'}
    MIN_YEAR = 2024
    MAX_YEAR = 2030