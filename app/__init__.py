from flask import Flask
import os
from app.security import BasicSecurity

def create_app():
    # Get the absolute path to the root directory
    root_dir = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
    
    # Set up template and static directories
    template_dir = os.path.join(root_dir, 'templates')
    static_dir = os.path.join(root_dir, 'static')
    
    # Create Flask app with configured directories
    app = Flask(__name__, 
                template_folder=template_dir,
                static_folder=static_dir,
                static_url_path='/static')

    # Add basic security
    BasicSecurity(app)
    
    # Import and register blueprints/routes
    from app.routes import scheduler_bp
    app.register_blueprint(scheduler_bp)

    return app