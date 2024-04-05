This folder contains Python scripts to generate a summary spreadsheet for the
results of Soundcheck tests for:
 - Headsets and earpeices MFP126, MFP166 and MFP270
 - SP1R and the legacy SmartPlugR1
 
To generate the summary run a Windows Command Prompt and navigate to this
folder.

Type 'py ./process_SP1R.py' or 'py ./process_126_166_270.py' to start the
summary generation for the appropriate product.

The script will prompt you for the full or relative path to the folder containing
the individual unit test results. I find it easy to have the folder open in 
Windows Explorer and to then copy the path from the folder view.

The 'raw' results are in .xls format, and the tool converts these into .xlsx format
before processing them.

The output is the file 'summary.xlsx' in the folder that contained the unit results.
