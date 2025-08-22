# Inventory Scripts

This repository contains all the Python tools and automation required to maintain a clean, streamlined inventory system for R&D chemicals and fluids. The solution is centered around an Excel file called `rnd_chemical_inventory.xlsx`, and the codebase automates inventory updates, stock taking, barcode generation, reorder notifications, and expired item cleanup.

## Repository Contents

There are five primary scripts in this repository:

### **combined_inventory_script.py** 
- Option 1 is to scan barcodes and update the InventoryData sheet in the Excel file
  - Prompts you for chemical details if the barcode is new, or usage amount if it's already in the system
- Option 2 is to take stock and/or mark empty bottles
- Option 3 is to exit the script

### **teamsmessage.py**
- Sends automated Teams alerts on 2 different things: an update on the status of critical chemicals, and a general summary table of anything expired, empty, unusable and things that are on the verge of being

### **reorder.py**
- Run when you want to check what needs to be ordered/ you already ordered something and want to update inventory (make sure it's excel on computer!)
- Automatically tracks what needs to be reordered and can remove expired/empty items from the sheet
- Once you set the date, items move to the OrderTracking sheet and off the main sheets

### **barcodegeneration.py**
- Generates barcodes and creates a printable PDF
- Outputs barcode images to the barcode_images folder

### **deletion_crontab.py**
- Moves items that have been reordered for their lead time + 1 month off of the ordertracking sheet onto the DeletedItems sheet

## Excel Workbook Structure

The main Excel file, `rnd_chemical_inventory.xlsx`, lives in OneDrive and contains 8 sheets:

- **InventoryData**: Where all the data is collected when a barcode is scanned or manually input
- **ReorderUpdates**: Entirely formulaic, based off InventoryData and ManufacturerLeadTimes to give accurate output of when things need to be reordered, lead times, emptiness and expiry
- **ManufacturerLeadTimes**: Table of approximate lead times for certain manufacturers to prevent extended wait periods
- **Audit_Log**: Keeps track of who uses the combined_inventory_script code and makes a log of the user and their general changes
- **GeneratedBarcodes**: Keeps track of all barcodes that have been created (don't delete anything on this sheet)
- **OrderTracking**: Where rows are stored after the reorder script is run and rows have been set to reorder
- **DeletedItems**: Where rows are stored after manufacturer lead time + a month runs out, and the item is moved here for tracking
- **StockTakingResults**: Stores detailed information about what was found in the scanner and what wasn't, more detailed than the summary

---

# Step-by-Step User Guide

## Scenario 1: You Need New QR Codes

1. Open the `barcodegeneration.py` code
2. Change the code at the top that looks for the excel file to match your file path
3. Run the code
4. Look up the PDF it creates in the search bar and print the barcodes on a sticker sheet
5. Apply the stickers to the necessary bottles

## Scenario 2: Taking Inventory or Stock (Single-User System)

### Initial Setup:
1. **Plug in the USB barcode scanner** to your computer
2. **Download the Excel file**: Go to OneDrive and download the most recent version of `rnd_chemical_inventory.xlsx` to your computer (**NOT online Excel**)
3. **Update file paths**: Get the code from GitHub and use `Ctrl+F` (or `Cmd+F` on Mac) to search for `'wb ='` and change the necessary path files that appear (there should be 2)
4. **Keep the Excel file closed** when running the code

### Running the Script:
When you run the script, you'll get a pop-up asking you to input:
- **1** = take/update inventory
- **2** = take stock/mark empty bottles  
- **3** = exit

### Option 1: Taking/Updating Inventory

**For New Items:**
1. Scan the barcode
2. Follow the prompts to fill out the necessary fields
3. Press OK to save it
4. Repeat as necessary

**For Updating Existing Items:**
1. Scan the barcode of an existing item
2. You'll get prompts about:
   - Date opened
   - Who opened it
   - How much fluid you used
3. Press OK to save
4. Repeat as necessary (now it will only promt you for how much fluid you used)

### Option 2: Stock Taking

You can do one of two things but they aren't explicitly separated, so you can do both at the same time if needed:

**Option A - Empty Bottles Only:**
1. Scan all the empty bottles
2. Press "finish stock taking"
3. Mark them all as empty
4. Close the window, then confirm the choice (another pop-up will appear)
5. open the excel file and save, then re-upload back to Onedrive

**Option B - Full Stock Check:**
1. Scan all the stock, click finish stock taking
2. Mark any bottles that happen to be empty during scanning and close the window by pressing x in the top left
3. After the pop-up, it will show you a summary
4. For further detail, open the workbook and check the **StockTakingResults** sheet, which keeps a log of what wasn't scanned
5. save the excel file and re-upload back to onedrive

### Finishing Up:
1. After scanning, press "yes" to terminate the Python window
2. On pycharm this works, but on sypder this has trouble so:
     If it doesn't close properly:
     - **Mac**: `Cmd + Option + Esc`
     - **Windows**: `Alt + F4`
4. **Save and upload**: Open the Excel file, save it (`Cmd+S` on Mac), and upload it back to OneDrive by clicking "Share" to keep it as the most updated copy

## Scenario 3: Setting Up Teams Alerts and Automatic Deletion (One Computer Setup)

### Initial Setup:
1. **Download the automation scripts**: Get `teamsmessage.py` and `deletion.py`
2. **Update file paths**: Change the paths for the Excel files at the top of each code file to match where you saved the Excel sheet on your computer
3. **Set up task scheduling**: Follow the crontab/task scheduler instructions below

### Scheduling Setup:
1. **Download timing**: At your most convenient time, download the `rnd_chemical_inventory.xlsx` sheet and save it over your previous copy
2. **Schedule the scripts**:
   - Teams messages and deletion: Once a week, but run deletion before the teams message
  
3. **File maintenance**: A few times a month, re-upload the file (Share → OneDrive) - download at night and re-upload the cleaned version in the morning to ensure deleted items are removed for everyone (it won't affect the main pages because the reordering script does that)

## Scenario 4: Reordering -> YOU HAVE TO RUN THIS CODE ACTIVELY, NOT IN THE BACKGROUND

The reorder code automatically:
- Keeps track of what needs to be reordered
- Removes expired/empty rows from the sheet by moving them to OrderTracking
- Once you set the reorder date, items move to the OrderTracking sheet and off the main sheets
- This is the only way to remove items off the main sheet, so make sure you run this code
- set it to run when you're at work
- 

---

# Technical Setup Instructions

## Prerequisites

- Python environment (PyCharm and Spyder were used for development)
- USB barcode scanner
- Access to OneDrive with the `rnd_chemical_inventory.xlsx` file

## Code Setup

1. **Update file paths**: For all scripts, change the file paths at the top of each code file to match your specific computer setup
2. **Install packages**: If packages are missing, install them using:
   ```bash
   pip install [package_name]
   ```
3. **Find path files**: Use `Ctrl+F` (or `Cmd+F`) and search for `"wb ="` to locate where file paths need to be updated

## Excel Formula Troubleshooting

If the ReorderUpdates sheet isn't updating properly with the InventoryData sheet, use these formulas (enter in the first cell of each column, then drag down to fill the entire column):

### ReorderUpdates Sheet Formulas:

**Column A2 (Chemical Name):**

=IF(InventoryData!B2="", "", InventoryData!B2)


**Column B2 (Bottle Nominal Volume (L)):**

=IF(InventoryData!F2="", "", InventoryData!F2)


**Column C2 (Remaining Quantity):**

=IF(InventoryData!G2="", "", InventoryData!G2)


**Column D2 (Expiry Date):**

=IF(InventoryData!I2="", "", InventoryData!I2)


**Column E2 (Expiry Alert):**

=IFS(ReorderUpdates!D2="Not Written", "Unknown", (ReorderUpdates!D2-TODAY())<0, "Expired",(ReorderUpdates!D2-TODAY())<VLOOKUP(G2,ManufacturerLeadTimes!A:B,2,FALSE), "Order More",(ReorderUpdates!D2-TODAY())>=VLOOKUP(G2,ManufacturerLeadTimes!A:B,2,FALSE), "Sufficient time")


**Column F2 (Quantity Alert):**

=IFS(
    ReorderUpdates!C2 = 0, "Empty",
    ReorderUpdates!C2< (ReorderUpdates!B2*0.1), "Do Not Use",
    ReorderUpdates!C2 < (ReorderUpdates!B2*0.2), "Order More",
    ReorderUpdates!C2 >= (ReorderUpdates!B2*0.2), "Sufficient Amount"
)


**Column G2 (Manufacturer):**
=IF(InventoryData!H2="", "", InventoryData!H2)


---

# Automated Script Installation

## For Mac (via crontab)

1. Open Terminal and run: `crontab -e`
2. Press `i` to insert, then paste the following lines (edit file paths and Python interpreter as needed):

```bash
# Weekly critical chemicals alert (Mondays at 3:02 AM)
2 3 * * 1 /opt/anaconda3/bin/python3 /Users/ishanasahota/Desktop/fluidAI/teamsmessage.py >> /Users/ishanasahota/Desktop/fluidAI/cronlog.txt 2>&1

# Daily reorder alerts (Daily at 3:02 AM)  
2 3 * * * /opt/anaconda3/bin/python3 /Users/ishanasahota/Desktop/fluidAI/reorder.py >> /Users/ishanasahota/Desktop/fluidAI/cronlog.txt 2>&1

# Daily deletion cleanup (Daily at 2:02 AM)
2 2 * * * /opt/anaconda3/bin/python3 /Users/ishanasahota/Desktop/fluidAI/deletion.py >> /Users/ishanasahota/Desktop/fluidAI/cronlog.txt 2>&1
```

3. Press `Esc`, then type `:wq` and hit `Enter` to save and exit
4. Verify setup with: `crontab -l`
5. Check that your Teams receives alerts at the correct times

## For Windows (via Task Scheduler)

1. **Open Task Scheduler**
2. **Create a new folder** (e.g., "MyTasks") to organize your scripts
3. **Create new tasks** with the following settings:
   - **Action**: Start a program
   - **Program/script**: Your Python executable path
   - **Add arguments**: Full path to your script (e.g., `reorderchemicalsfluidai.py`)

4. **Schedule (example)**:
   - **Reorder script**: Bi-Weekly at 9 AM (When you're in work/ only run when you're updating ordering status/ need to remove exmpty/expired rows from the main script)
   - **Weekly alerts**: Weekly at 9 PM Monday  
   - **Auto-delete**: Daily at 8 PM

---

# System Logic & Automation

## How the ReorderUpdates Sheet Works
The ReorderUpdates sheet is formula-based, pulling data from InventoryData and calculating:
- **Quantity Alert**: Based on remaining quantity vs. bottle size
- **Expiry Alert**: Based on expiry date vs. manufacturer lead times
- These alerts are automatically picked up by the reorder and weekly alert scripts


## OneDrive Best Practices

**⚠️ CRITICAL**: Never edit ReorderUpdates unless it's to fix the formulas (Hint: if you run the reorderalerts, the code has all the formulas to automatically repopulate it)

**Always follow this workflow**:
1. Download the most recent version from OneDrive to your desktop
2. Run your updates or scripts with the file closed
3. Save the file after making changes
4. Upload the updated version back to OneDrive using "Share"

**For the person running crontab/ task scheduler**: 
- Download the newest version of the Excel file from OneDrive daily for accurate automation
- Set calendar reminders to download the file at a consistent time
- A few times per month, re-upload the cleaned version to OneDrive to remove deleted items for all users

---

# Important Notes

- The system requires one person to manage the crontab/task scheduler setup
- All file paths in the scripts must be updated to match your specific computer setup
- The barcode scanner should be plugged in before running inventory scripts
- Keep the Excel file closed when running Python scripts
- Always save and re-upload to OneDrive after making changes to keep the system synchronized
