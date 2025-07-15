# Combine BOM
This repository is home to the python script that will combine multiple BOMs into one large concise BOM, which alleviates the headache associated with purchasing materials. This script, much like the original, is intended to work with .xlsx files, opposed to .xls.  

# Description 
When running this script, the user will be able to create one large flat BOM for purchasing purposes.  This script will only operate on .xlsx files, and not .xls files. This script will automatically sift through all files in the current working directory, and with each file, it will iterate over all sheets.  To exclude a file from processing, simply and temporarily change the file extension to something other than _xlsx_.  

Each BOM _must_ contain headings: __QPN__ | __QTY__ | __UOM__ |__DES__ | __REF__ | __MFG__ | __MFGPN__ | __CR1__ | __CR1PN__ | __NOTES__ 

Subtle discrepancies will be accepted.  For example, _Reference_ or _Ref_, will be accepted for __REF__.  Since the application automatically locates the location of various data columns, it needs to seek out this header before starting. Data within the column can be completely blank, so long as the heading title is in place.  

# Revisions
v1.0 -- Initial (tested) release.  It's worth noting that this version works with unicode, and so special symbols, if encountered, shouldn't cause a crash.  

v1.1 -- No longer are internal white spaces removed from descriptions, notes, etc., but rather only those that are leading or trailing.  This prevents descriptions, notes, etc. from being run together.

v1.2 -- Bug addressed which is exposed if workbook contains a revision page.  Just like the .XLS parser, this application incorporates the "sheet_valid" flag, and thus won't processes the BOM unless it has been determined that the sheet is valid.  The issue related to opened file types has been corrected.  Now, only .xlsx files are opened and parsed.  

v2.0 -- Ported for Python 3.8+.  Renamed the file.  Tested with BOMs, and appears to function well. 

v2.1 -- leading b' text now removed from cell values thus allowing for proper BOM generation.  Additional _unicode_ errors found and corrected by way of using _encode_ method. 

v2.2 -- Fixed minus-one bug in which not all files within a directory were being  were being iterated over.  A component reference field is now required in the heading.  

v2.3 -- Added UOM column.  Searching for data headers is now done through regular expressions.  

v2.4 -- Greatly increased logging capabilities so one can more-easily determine if a BOM file is improperly formatted.  

v2.5 -- Fixed issue in which the header was being prematurely detected due to text containing keywords before the actual header was processed.  

v2.6 -- Increased logging capability so it is easier to identify BOM formatting issues.  System will now exit upon detecting a BOM format issue, but will first alert user of problem.  Said format issue is also logged via logging.log.  

v2.7 -- Modified mechanism that determines if REF column is in the BOM.  The NOTES RE was expanded to include both NOTES and NOTE.

v2.8 -- This README file was updated to indicate that we _must_ have a heading that represents UOM.  Also, if a consolidated BOM already exists, then it will first be deleted before additional processing can occur. 