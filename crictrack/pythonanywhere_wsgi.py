# Paste this into the WSGI configuration file on PythonAnywhere (Web tab -> WSGI configuration file).
# Replace YOURUSERNAME. Turn on "Force HTTPS" on the Web tab.
import os, sys
path = "/home/YOURUSERNAME/cricket-tracker"
if path not in sys.path:
    sys.path.insert(0, path)
os.environ["SECURE_COOKIES"] = "1"      # login cookie only over https
from app import app as application, _load
_load(force=True)
