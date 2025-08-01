This repository contains all the Python tools and automation required to maintain a clean, streamlined inventory system for R&D chemicals and fluids. The solution is centered around an Excel file called rnd_chemical_inventory.xlsx, and the codebase automates inventory updates, barcode generation, reorder notifications, and expired item cleanup.

There are five primary scripts in this repository:
  take_inventory.py

-> Used to scan barcodes and update the InventoryData sheet in the Excel file.


->Prompts you for chemical details if the barcode is new, or usage amount if it's already in the system.


  weeklycriticalchemicalsfluidai.py

->Sends automated Teams alerts every Monday at 9 AM for chemicals that are critically low or expired.


  reorderchemicalsfluidai.py
->Sends daily alerts for any chemicals that need reordering or have upcoming expiry.


  barcodegeneration.py

->Generates barcodes and creates a printable PDF.

->Outputs barcode images to the barcode_images folder.


  deletion_crontab.py

->Automatically deletes any items in InventoryData that have either:
->Expired more than 30 days ago, or
->Have zero remaining quantity for more than 30 days.
->Moved deleted items to a DeletedItems sheet for record-keeping.


**How It Works**
The main Excel file, rnd_chemical_inventory.xlsx, lives in OneDrive and should always be downloaded fresh to excel on Mac before use.
The primary data lives in the InventoryData sheet.

IF YOU ARE TAKING INVENTORY:

When runtotakeinventory.py is run:
It scans or creates entries based on barcode input.
If it's a new chemical, you’ll be prompted to enter:
Chemical Name
Lot Number
Nominal Volume
Manufacturer
Expiry Date
All updates are automatically saved to the Excel file. After your finished, open the excel file and command + S the excel worksheet InventoryData and in the top right, where it says share, press that and upload it back to OneDive, therefore keeping the excel sheet as updated as possible


The ReorderUpdates sheet is formula-based, pulling data from InventoryData and calculating:
Quantity Alert
Expiry Alert
These alerts are then picked up by the reorder and weeklyalerts scripts to notify teams.
In order for this to work accurately, the person who has the crontab running the 3 scripts must download the newest version of the excel file from OneDrive everyday. (I couldn't find a better solution, but if any of the devs or actual coders have any suggestions for this, please fix this)

**deletion_crontab Logic**


This script performs automated cleanup of expired or depleted items:


Runs daily (e.g., at 2 AM via crontab or Task Scheduler).


Only deletes items 30+ days after expiry or emptiness, which means it will still send alerts from when it's originally expired, giving us enough time to reorder, and then after a month, is deleted from the sheet and moved to DeletedItems (for reference).


Deleted items are archived in a DeletedItems sheet with a timestamp.


You do not need to run a separate deletion for ReorderUpdates — because that sheet is formula-based, any deletions in InventoryData automatically clear the corresponding alert rows.


**NOTE**:   As many times as you see necessary during the month, upload it twice a day (download it once, possibly at night, to have the code run, and the next morning, upload it to make sure all the deleted items have been removed), press share, and share the excel doc, now free of any deleted items, back to onedrive. the 2- uploads don't have to be as daily as the nightly download, because if you download it once, the deletion code will still remove the items that are expired and past 30 days again and again, meaning your alerts will still be accurate, but for the convenience of everyone else, it won't be displayed. So a few times a month, upload back the cleaned copy!


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
2 3 * * 1 /opt/anaconda3/bin/python3 /Users/ishanasahota/Desktop/fluidAI/weeklycriticalchemicalsfluidai.py >> /Users/ishanasahota/Desktop/fluidAI/cronlog.txt 2>&1


2 3 * * * /opt/anaconda3/bin/python3 /Users/ishanasahota/Desktop/fluidAI/reorderchemicalsfluidai.py >> /Users/ishanasahota/Desktop/fluidAI/cronlog.txt 2>&1


2 2 * * * /opt/anaconda3/bin/python3 /Users/ishanasahota/Desktop/fluidAI/deletion_crontab.py >> /Users/ishanasahota/Desktop/fluidAI/cronlog.txt 2>&1


Press Esc, then type :wq and hit Enter to save and exit.
Verify setup with: crontab -l
Check that your Teams receives alerts at the correct times.


For Windows (via Task Scheduler)
Open Task Scheduler.
Create a new folder (e.g., "MyTasks") to organize your scripts.
Create new tasks with the following settings:
Action: Start a program
Program/script: Your Python executable path
Add arguments: Full path to your script (e.g., reorderchemicalsfluidai.py)
Schedule:
Reorder script: Daily at 9 AM
Weekly alerts: Weekly at 9 AM Monday
Auto-delete: Daily at 9 AM

**For All This Code to Work**
->The only and the most important thing to remember is to download the excel file to the mac, save it and re-upload it to onecrive once you've made edits
->If you have the crontab, set reminders on calendar or wherever else to download the file at a certain time everyday, and edit the crontab hour (not the 0 or the spacing, but the number) if you want the code to run through the most updated version (aka right after you download it

