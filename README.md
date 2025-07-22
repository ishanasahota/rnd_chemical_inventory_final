In this repository is my work towards making R&D an easy-to-use inventory system. There are 4 pieces of code in here that do four different things. 2 of them, 'weeklycriticalchemicalsfluidai.py'
and 'reorderchemicalsfluidai.py' are responsible for reorder updates that are sent to teams. 'runtotakeinventory.py' is self-explanitory and 'finalbarcodegeneration.py' is used to generate 
barcodes. 

The process works like this. There is an excel file called 'rnd_chemical_inventory' where the fluid inventory is stored. The first sheet, InventoryData, is where, when 'runtotakeinventory.py'
is run, the information is stored. 'ReorderUpdates' takes four columns of InventoryData and then adds two of it's own, Expiry Alerts and Quantity Alert. This sheet, ReorderUpdates, and those 
two columns are what 'weeklycriticalchemicalsfluidai.py' and 'reorderchemicalsfluidai.py read and send Teams notifications on. When you run out of barcodes, there's 'finalbarcodegeneration.py'.
Before running, look up the folder barcode_images with your F4 key (if you haven't run it before it won't exist) and clear it's contents BUT not the folder. This will make sure only the newly
generated barcodes end up on the PDF and not reprints of old ones (Which will mess everything up). Then once you stick a barcode on, you can add information to the sheet. Go onto OneDrive and
download to desktop Excel the most recent copy of rnd_chemical_inventory. Once that is loaded, make sure the 'runtotakeinventory.py' has the correct path at the top of the code to access the 
excel file. That is the only thing you need to change. Then, making sure that the excel file is closed (and not minimized), run the code. There will be a pop-up that asks you to scan the 
barcode, and do as it says. If it has been scanned before, it will prompt you to input how much you used. If it has never been scanned before, it will prompt you for the Chemical Name, 
Lot Number, Nominal Amount, Manufacturer and Expiry Date. Input all of those, and it will save automatically to the excel file. Once you are done scanning, click 'yes' to exit scanning, 
and if the window doesn't close, press command + option + esc to close the python tab. Then open the excel file, save it, and press 'Share' at the top and upload the new, updated copy back
to OneDrive. 

Never open the excel file off of your desktop, only download it from OneDrive to ensure you have the newest copy


If you need to run 'weeklycriticalchemicalsfluidai.py' and 'reorderchemicalsfluidai.py', you need to set up either crontab or task scheduler,
depending on if you have a Windows or a Mac. 

For Mac, go onto your terminal and type in 'crontab -e' (but with no apostrophes).
then, once your inside, press 'i' to insert the following 2 lines
0 9 * * 1 /opt/anaconda3/bin/python3 /Users/ishanasahota/Desktop/fluidAI/weeklycriticalchemicalsfluidai.py >> /Users/ishanasahota/Desktop/fluidAI/cronlog.txt 2>&1
0 9 * * * /opt/anaconda3/bin/python3 /Users/ishanasahota/Desktop/fluidAI/reorderchemicalsfluidai.py >> /Users/ishanasahota/Desktop/fluidAI/cronlog.txt 2>&1
 change these to your specific path file where the code is located on your computer, BUT do not mess up the spacing, as the cron job will not run. Once the information is pasted, click esc, and then type ':wq' which will exit you from the crontab. To see if they properly pasted, check by typing 'crontab -l' into your terminal. if it displays what you see above, then it's been inputted correctly. Then check your Teams when 9 AM rolls around to see if it worked.

 For Windows, you have to set up task scheduler.
