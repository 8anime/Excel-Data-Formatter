
import os

from dataConverter.textToNumeric import fixExcelNumberFormatting

INPUT_FILE = os.path.join('inputFolder', 'Tourism dataset.xlsx')       # Path to the input Excel file containing the tourism dataset.
OUTPUT_FILE = os.path.join('outputFolder', 'Tourism dataset.xlsx')     # Path to save the corrected Excel file after processing.
# Name of the sheet within the Excel file that contains the tourism data.
SHEET_NAME = 'tourism data'  # Update this to change the sheet name in the Excel file
# The delimiter character to be removed from text-formatted numbers (e.g. ',' for '1,000').
DELIMITER = ',' 
# A list of column letters to be widened in the Excel sheet. These columns will have their widths adjusted.
COLUMN_TO_WIDEN = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']

if __name__ == '__main__':
    fixExcelNumberFormatting(INPUT_FILE, OUTPUT_FILE, SHEET_NAME, DELIMITER, COLUMN_TO_WIDEN)


