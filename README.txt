================================================
  THE AWARENESS ENGINE — Setup & Run Instructions
================================================

WHAT THIS DOES
--------------
You upload your .xlsx spreadsheet. The app reads your
JAN, FEB, and MAR data, shows you each spending category,
lets you label each one as Necessities / Luxuries / Future Self,
then generates a clean downloadable spreadsheet with
your real percentages.


REQUIREMENTS
------------
You need Python installed on your computer.
Don't have it? Download it free at: https://www.python.org/downloads/
(Download the version that says "3.x" — e.g. Python 3.11)

During install on Windows: CHECK the box that says
"Add Python to PATH" — very important!


STEP 1 — Open your Terminal (Command Prompt)
--------------------------------------------
Windows: Press Windows key, type "cmd", press Enter
Mac:     Press Cmd + Space, type "Terminal", press Enter


STEP 2 — Navigate to this folder
---------------------------------
Type this command and press Enter:

  cd path/to/awareness_engine

Example on Windows:  cd C:\Users\YourName\Downloads\awareness_engine
Example on Mac:      cd ~/Downloads/awareness_engine


STEP 3 — Install the required packages (one time only)
-------------------------------------------------------
Type this and press Enter:

  pip install -r requirements.txt

Wait for it to finish. You'll see some text scroll by — that's normal.


STEP 4 — Start the app
-----------------------
Type this and press Enter:

  python app.py

You should see something like:
  * Running on http://127.0.0.1:5050


STEP 5 — Open in your browser
------------------------------
Open Chrome, Firefox, or Edge and go to:

  http://127.0.0.1:5050

The app will load and guide you through the 3 steps.


STEP 6 — Use the app
---------------------
1. Upload your .xlsx file (the one with JAN, FEB, MAR tabs)
2. For each spending category, click the bucket it belongs in:
     Necessities | Luxuries | Future Self
3. Click "Generate My Awareness Report"
4. Click the green Download button
5. Open the downloaded file in Excel or Google Sheets


STOPPING THE APP
----------------
Go back to your terminal and press:  Ctrl + C


STARTING AGAIN NEXT TIME
-------------------------
Just repeat Step 4:  python app.py


TROUBLESHOOTING
---------------
"pip is not recognized" →
  Try: python -m pip install -r requirements.txt

"python is not recognized" →
  Python might not be in your PATH. Reinstall Python and
  make sure to check "Add Python to PATH" during setup.

"Port already in use" →
  Another app is using port 5050. Open app.py, find the last line
  and change 5050 to 5051 (or any number), then restart.

App shows $0.00 for everything →
  Your spreadsheet's TOTALS column might be in a different spot.
  Make sure the column that says "TOTALS" is present in row 1
  of each month tab (JAN, FEB, MAR).
================================================
