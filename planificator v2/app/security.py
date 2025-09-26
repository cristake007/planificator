from flask import request, abort

class BasicSecurity:
    def __init__(self, app):
        self.app = app
        self.setup_security()

    def setup_security(self):
        @self.app.before_request
        def check_request():
            # Skip checks for static files
            if request.path.startswith('/static/'):
                return
                
            # Check for suspicious patterns in raw data
            raw_data = request.get_data()
            bad_patterns = [b'\x03\x00\x00/', b'mstshash']
            
            if isinstance(raw_data, bytes):
                if any(pattern in raw_data for pattern in bad_patterns):
                    return abort(403)
            
            # Check for suspicious headers
            for header in request.headers.values():
                if any(ord(c) < 32 for c in str(header)):
                    return abort(403)

        @self.app.after_request
        def add_basic_security(response):
            response.headers['X-Content-Type-Options'] = 'nosniff'
            response.headers['X-Frame-Options'] = 'SAMEORIGIN'
            return response