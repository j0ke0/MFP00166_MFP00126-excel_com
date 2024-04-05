# Process_126_166_270
# Simple script to ask for a path to a folder with test results for
# a bunch of MFP126/166/270 products.
# First we call the script to convert all .xls files to .xlxs because
# the funky python excel library only handles .xls.
# The we call the script that generates the summary .xlxs files
import convert_xls2xlsx
import create_data_summary_MFP126_MFP166_MFP270


if __name__ == "__main__":
    # Assuming the text file contains the path
    with open('location.txt', 'r') as file:
        path = file.readline().strip()
    
    convert_xls2xlsx.xls_2_xlsx(path)
    create_data_summary_MFP126_MFP166_MFP270.processResults(path)