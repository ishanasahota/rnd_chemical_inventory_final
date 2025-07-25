This repository contains all the Python tools and automation required to maintain a clean, streamlined inventory system for R&D chemicals and fluids. The solution is centered around an Excel file called rnd_chemical_inventory.xlsx, and the codebase automates inventory updates, barcode generation, reorder notifications, and expired item cleanup.

There are four primary scripts in this repository:
runtotakeinventory.py
Used to scan barcodes and update the InventoryData sheet in the Excel file.
Prompts you for chemical details if the barcode is new, or usage amount if it's already in the system.


weeklycriticalchemicalsfluidai.py
Sends automated Teams alerts every Monday at 9 AM for chemicals that are critically low or expired.


reorderchemicalsfluidai.py
Sends daily alerts for any chemicals that need reordering or have upcoming expiry.


finalbarcodegeneration.py
Generates barcodes and creates a printable PDF.
Outputs barcode images to the barcode_images folder.


auto_delete_expired.py
Automatically deletes any items in InventoryData that have either:
Expired more than 30 days ago, or
Have zero remaining quantity for more than 30 days.
Moved deleted items to a DeletedItems sheet for record-keeping.

**How It Works**
The main Excel file, rnd_chemical_inventory.xlsx, lives in OneDrive and should always be downloaded fresh before use.
The primary data lives in the InventoryData sheet.


When runtotakeinventory.py is run:
It scans or creates entries based on barcode input.
If it's a new chemical, you’ll be prompted to enter:
Chemical Name
Lot Number
Nominal Volume
Manufacturer
Expiry Date
All updates are automatically saved to the Excel file.


The ReorderUpdates sheet is formula-based, pulling data from InventoryData and calculating:
Quantity Alert
Expiry Alert
These alerts are then picked up by the reorder scripts to notify teams.

Barcode Generation Notes:
Before running finalbarcodegeneration.py:
Go to the barcode_images folder (press F4 in your Finder to locate it).
Clear its contents (but do not delete the folder itself).
This ensures the PDF contains only the newly generated barcodes.

Deletion Logic
This script performs automated cleanup of expired or depleted items:
Runs daily (e.g., at 2 AM via crontab or Task Scheduler).
Only deletes items 30+ days after expiry or depletion.
Items are not deleted immediately so alerts still have time to notify teams.
Deleted items are archived in a DeletedItems sheet with a timestamp.
You do not need to run a separate deletion for ReorderUpdates — because that sheet is formula-based, any deletions in InventoryData automatically clear the corresponding alert rows.

OneDrive and best Practices
Never edit the Excel file directly on OneDrive.
Instead, always:
Download the most recent version to your desktop.
Run your updates or scripts.
Save and upload the updated version back to OneDrive.

CRONTAB & DAILY SCRIPT INSTALLATION
For Mac (via crontab)


Open Terminal and run: crontab -e
Press i to insert, then paste the following lines (edit file paths and Python interpreter as needed):
0 9 * * 1 /opt/anaconda3/bin/python3 /Users/ishanasahota/Desktop/fluidAI/weeklycriticalchemicalsfluidai.py >> /Users/ishanasahota/Desktop/fluidAI/cronlog.txt 2>&1
0 9 * * * /opt/anaconda3/bin/python3 /Users/ishanasahota/Desktop/fluidAI/reorderchemicalsfluidai.py >> /Users/ishanasahota/Desktop/fluidAI/cronlog.txt 2>&1
0 2 * * * /opt/anaconda3/bin/python3 /Users/ishanasahota/Desktop/fluidAI/auto_delete_expired.py >> /Users/ishanasahota/Desktop/fluidAI/cronlog.txt 2>&1
Press Esc, then type :wq and hit Enter to save and exit.
Verify setup with: crontab -l
Check that your Teams receives alerts at the correct times.


✅ For Windows (via Task Scheduler)
Open Task Scheduler.
Create a new folder (e.g., "MyTasks") to organize your scripts.
Create new tasks with the following settings:
Action: Start a program
Program/script: Your Python executable path
Add arguments: Full path to your script (e.g., reorderchemicalsfluidai.py)
Schedule:
Reorder script: Daily at 9 AM
Weekly alerts: Weekly at 9 AM Monday
Auto-delete: Daily at 2 AM

**Tips for Daily Use**
Ensure you have the latest copy of the Excel file downloaded each day.
At the end of the day, upload the updated Excel file back to OneDrive to keep everything synced.
Set calendar alerts for yourself until this becomes a habit.
