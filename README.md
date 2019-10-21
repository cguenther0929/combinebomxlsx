# Combine BOM
This repository is home to the python script that will combine multiple BOMs into one large concise BOM, which alleviates the headache associated with purchasing materials. This script, much like the original, is intended to work with .xlsx file, opposed to .xls.  

# Description 
When running this script, the user will be able to create one large flat BOM for purchasing purposes.  This script will only operate on .xlsx files, and not .xls files. This script will automatically sift through all files in the current working directory, and with each file, it will iterate over all sheets.  If the user is wanting to skip a file, he/she could simply change the extension of the file temporarily to something other than .xlsx.   

# Revisions
v1.0 -- Initial (tested) release.  It's worth noting that this version works with unicode, and so special symbols, if encountered, shouldn't cause a crash.  

v1.1 -- No longer are internal white spaces removed from descriptions, notes, etc., but rather only those that are leading or trailing.  This prevents descriptions, notes, etc. from being run together.

v1.2 -- Bug addressed which is exposed if workbook contains a revision page.  Just like the .XLS parser, this application incorporates the "sheet_valid" flag, and thus won't processes the BOM unless it has been determined that the sheet is valid.  The issue related to opened file types has been corrected.  Now, only .xlsx files are opened and parsed.  