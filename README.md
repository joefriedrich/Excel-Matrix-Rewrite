# Excel-Matrix-Rewrite (aka pySAP-Security)
Code written to replace VBA code.  Looks up items in excel based on user input.
# Requirements
Requires Python 3.5 and openpyxl 2.3.2.  Other versions may work, but this was our development environment.  
It is not currently Python 2 or openpyxl 2.4+ compatable.
# Instructions
Move the sample data files (.xlsx) where you want, but you must document the location in the main code.
There is a data structure called a 'Company', and this structure must contain this information.
If you want to work with your own data, you must maintain this.
# Future Plans
0. Make code PEP-8 compliant. 
1. Re-write single-approver functionality to move data out of code.
2. Make a .ini file to house company specific information (moving data out of code).
3. Migrate away from Excel Spreadsheets to .csv files (which can still be read by Excel).
4. Automate updates to .csv files from the source.
5. GUI!! (in Flask?)
