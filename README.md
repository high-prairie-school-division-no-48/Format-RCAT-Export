# Format-RCAT-Export

This project aims to modify a raw RCAT export provided by Edubest to match the formatting required by the RCAT import function in Dossier.

Important: only screening test results are imported into Dossier

![image](https://github.com/high-prairie-school-division-no-48/Format-RCAT-Export/assets/87395701/c2337f00-19ba-428b-a46f-dfbf0916a6dd)

## Main Components
1. formatRCATExport.py - Python script that modifies input file to match Dossier's RCAT import function
2. ######.xlsx - Excel spreadsheet used by the Python script (raw RCAT export provided by Edubest)

## Basic Walkthrough
1. Save raw RCAT export file to your local disk
2. Modify the _inputFilePath_ variable in the script (line 11) to match the path of your saved RCAT file

![image](https://github.com/high-prairie-school-division-no-48/Format-RCAT-Export/assets/87395701/93bef0a1-a75d-47ac-8139-96c937a811e2)

3. Save changes to the script
4. Run script and the following actions will be performed automatically in this order:
   - Filter the rows to remove non-screening test results
   - Filter the columns to only include headers used by Dossier's import function
   - Rename the column headers to match the ones used by Dossier's import function
   - Reorder the columns to match the order used by Dossier's import function
   - Convert the Date column to use _yyyy-mm-dd_ formating 
   - Group rows by _Student → Passage → Skill Category_ to collect total correct and total number of questions per skill category
   - Calculate each skill category percentage as well as the overall percentage for each individual test
5. Output file will be generated in the same directory as the script

