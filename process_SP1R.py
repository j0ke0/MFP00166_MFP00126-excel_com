# Process_SPR
# Simple script to ask for a path to a folder with test results for
# a bunch of SPR products.
# First we call the script to convert all .xls files to .xlxs because
# the funky python excel library only handles .xls.
# The we call the script that generates the summary .xlxs files
import convert_xls2xlsx
import create_data_summary_SP1R_SP01

if __name__ == "__main__":
	path = input("Please provide the results path:\n> ")
	convert_xls2xlsx.xls_2_xlsx(path)
	create_data_summary_SP1R_SP01.processResults(path)