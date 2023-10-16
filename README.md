# Excel Data Processing Tool

The Excel Data Processing Tool is a Python script designed to streamline and automate the process of cleaning and formatting data in Excel files. It's especially useful for ensuring data quality, accuracy, and consistency when working with Excel datasets.

This tool has the following key features:
- Removes specified delimiters (e.g., commas) from text-formatted numbers in Excel cells.
- Converts text-formatted numbers into actual numeric values for accurate calculations and analysis.
- The tool allows you to specify columns that should be widened to ensure data is displayed correctly in Excel.

## Features

- **Delimiter Removal:** The tool can remove specified delimiters (e.g., commas) from text-formatted numbers in Excel cells.

- **Numeric Conversion:** It converts text-formatted numbers into actual numeric values for accurate calculations and analysis.

- **Column Widening:** The tool allows you to specify columns that should be widened to ensure data is displayed correctly in Excel.

## Prerequisites

- Python 3.x installed.

- The `openpyxl` library. You can install it using `pip install openpyxl`.

## Usage

1. Clone this repository or download the project folder.

2. Open the project folder in a text editor or integrated development environment (IDE).

3. Define the following parameters in the main.py file:

    - `inputFile`: Path to the input Excel file.
    
    - `outputFile`: Path to save the corrected Excel file.
    
    - `sheetName`: Name of the sheet within the Excel file.
    
    - `delimiter`: The delimiter character to be removed from text-formatted numbers (e.g., ',' for '1,000' or 1,333,399).
    
    - `columnToWiden`: A list of column letters to be widened (e.g., ['A', 'B']).

    NB// Make sure the file to be processed is in the "inputFolder" directory located in the root directory

4. Run the main.py script. 

5. The script will process the Excel file, remove delimiters, convert text-formatted numbers and widen columns as specified.

6. The corrected Excel file will be saved with the name you provided in `outputFile`.

7. The script will print a success message if changes were made, or it will indicate if no changes were needed.

## Example

Here's an example of how to use the script:

```python
# Define parameters
inputFile = 'input.xlsx'
outputFile = 'output.xlsx'
sheetName = 'Data'
delimiter = ','
columnToWiden = ['A', 'B']

# Run the script
fixExcelNumberFormatting(inputFile, outputFile, sheetName, delimiter, columnToWiden)
