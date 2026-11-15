# PythonAnywhere WSGI configuration example.
# Replace YOUR_PYTHONANYWHERE_USERNAME with your actual PythonAnywhere username.
import sys

project_home = "/home/YOUR_PYTHONANYWHERE_USERNAME/nusuk_code_submission_portal"
if project_home not in sys.path:
    sys.path.insert(0, project_home)

from app import app as application
