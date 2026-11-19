CricPay - Cricket tournament payment tracker
============================================

CricPay is a small web app for cricket teams to track tournament payments.
There is no database: an Excel workbook is the data store, and every change
made on the website is written straight back into that Excel file.


FEATURES
--------
- Tournaments: one sheet per tournament, created from the website
  (fee, matches, players per match, advance, expected cost).
- Fee-based costs: match cost = tournament fee / number of matches;
  each match's cost is shared by the players who actually played it.
  Example: Rs 42,000 / 6 matches = Rs 7,000 per match;
           12 players -> Rs 583 each, 10 players -> Rs 700 each.
- Played Yes/No toggles that recalculate everyone's amounts.
- Payments: record advance / match fee / other payments (UPI, cash, bank),
  with a full payment history and Undo.
- Organiser & kitty: organiser fee, paid to organiser, balance,
  still to collect, and kitty in hand.
- Custom amount per player and extra charges.
- Accounts and roles:
    Admin   - whole site, all tournaments, manages everyone
    Captain - read/write for their own tournaments, manages their players
    Player  - view only, their own tournaments
- Sheet view: edit any non-formula cell of the workbook in the browser.
- Automatic Excel backups, CSV export, mobile-friendly UI, IST time.


PROJECT LAYOUT
--------------
app.py                  Flask app: reads/writes the Excel file, all APIs
auth.py                 Sign-in, registration, roles, per-tournament access
manage_users.py         Command-line account recovery (reset password, make admin)
pythonanywhere_wsgi.py  Template WSGI file for PythonAnywhere
requirements.txt        Python packages
start-mac-linux.sh      Start locally on Mac/Linux
start-windows.bat       Start locally on Windows
start-public.sh         Start + share via a fixed ngrok address (Mac)
static/                 Web pages (ledger.html, login.html, index.html) and icons
data/                   NOT in git - the Excel file, accounts and backups live here


REQUIREMENTS
------------
- Python 3.10 or newer
- Packages in requirements.txt (Flask, openpyxl, pycel, waitress)


RUN IT LOCALLY
--------------
1. Clone the repository:
     git clone https://github.com/<your-username>/cricpay.git
     cd cricpay

2. Create a virtual environment and install packages:
     python3 -m venv .venv
     source .venv/bin/activate          (Windows: .venv\Scripts\activate)
     pip install -r requirements.txt

3. Put your Excel workbook in the data folder:
     mkdir data
     copy your file to data/2025_Cricket_Tournaments.xlsx
   (or point to any other file with:  export TRACKER_XLSX=/path/to/file.xlsx)

4. Start the app:
     python app.py
   Open http://localhost:8000 and create the first account - it becomes the admin.


DEPLOY ON PYTHONANYWHERE
------------------------
1. Bash console:
     git clone https://github.com/<your-username>/cricpay.git
     mkvirtualenv cricpay --python=python3.11
     pip install -r ~/cricpay/requirements.txt
2. Upload your Excel file to ~/cricpay/data/ (Files tab).
3. Web tab -> Add a new web app -> Manual configuration -> Python 3.11.
   Virtualenv: cricpay
4. Open the WSGI configuration file and use:

     import os, sys, time
     os.environ["TZ"] = "Asia/Kolkata"; time.tzset()
     os.environ["SECURE_COOKIES"] = "1"
     project = "/home/<your-username>/cricpay"
     if project not in sys.path:
         sys.path.insert(0, project)
     os.chdir(project)
     from app import app as application

5. Turn on "Force HTTPS", click Reload, open the site, create the admin account.

Updating later:
     cd ~/cricpay && git pull
   then click Reload on the Web tab. Your data folder is never touched by git.


SETTINGS (environment variables)
--------------------------------
TRACKER_XLSX     Excel file to use (default: data/2025_Cricket_Tournaments.xlsx)
PORT             Port for local runs (default: 8000)
SECURE_COOKIES   Set to 1 on an https site
EDIT_PIN         Optional; protects creation of the first admin account


ACCOUNT RECOVERY
----------------
From the project folder:
     python manage_users.py list
     python manage_users.py reset-password <username>
     python manage_users.py make-admin <username>
     python manage_users.py create-admin <username> "<Full name>"


DATA AND PRIVACY
----------------
The data/ folder holds players' names, payment records and account password
hashes. It is excluded by .gitignore and must never be pushed to GitHub.
Keep the repository private, and back up data/ separately
(the site's Excel button downloads the current workbook).