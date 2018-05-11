# Excel-Matrix-Rewrite (aka pySAP-Security)
Code written to replace VBA code.  Looks up items in CSVs based on user input.
# Requirements
Requires Python 3.5.  Other versions may work, but this was our development environment.  
It is not currently Python 2 compatable.
# Instructions
Move the sample data folders where you want, but you must document the location in the main code.
There is a data structure called a 'Company', and this structure must contain this information.
If you want to work with your own data, you must maintain this.
# Future Plans
0. Make code PEP-8 compliant. 
1. Re-write so regional is only prompted if the user picks a regional role.
2. Make a .ini file to house company specific information (moving data out of code).
3. Automate updates to .csv files from the source.
4. GUI!! (in Flask?)
