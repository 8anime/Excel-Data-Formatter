
import openpyxl


def fixExcelNumberFormatting(inputFile, outputFile, sheetName, delimiter, columnToWiden):
    """
    Process and format an Excel dataset with the following operations:
    1. Remove specified delimiters from text-formatted numbers and convert them to numeric values.
    2. Widen specified columns to ensure data is displayed correctly.

    Args:
        inputFile (str): Path to the input Excel file.
        outputFile (str): Path to save the corrected Excel file.
        sheetName (str): Name of the sheet within the Excel file.
        delimiter (str): The delimiter character to be removed from text-formatted numbers (e.g., ',' for '1,000').
        columnToWiden (list): A list of column letters (e.g., ['A', 'B']) to be widened.

    Returns:
        None

    Example:
        To process an Excel file 'input.xlsx', remove commas as delimiters, convert text-formatted numbers to actual numbers,
        and widen columns A and B in the 'Data' sheet to a width of 15:
        fixExcelNumberFormatting('input.xlsx', 'output.xlsx', 'Data', ',', ['A', 'B'])
    """
    try:
        # Load the Excel file
        wb = openpyxl.load_workbook(inputFile)
        # Specify the sheet you want to work with
        sheet = wb[sheetName]
        # Flag to track if any changes were made
        changesMade = False

        # Iterate through all rows and columns
        for row in sheet.iter_rows():         # Iterate through rows in the Excel file
            for cell in row:                  # In a row, iterate through the cell fields
                if cell.data_type == 's':     # Check if a value in the cell field is a text or string
                    # Remove the specified delimiter
                    cell.value = cell.value.replace(delimiter, '')  # Remove the delimiter(,) in the 'text integer' and replace it with an empty space
                    # Attempt to convert the content to a number
                    try:
                        cell.value = float(cell.value)              # Convert the value without the delimiter into a float value
                        changesMade = True                          # If process is successful update this with True
                    except ValueError:
                        # If conversion to a number fails, continue to the next cell
                        pass

        if changesMade:  # Checks if the process was successful
            # This function widens the column fields to ensure data is displayed correctly
            widenColumns(sheet, columnToWiden, width=15)  # Adjust the width as needed

            # Save the corrected Excel file to another new file(output file)
            wb.save(outputFile)
            print(f'Success: Delimiters removed, numbers formatted as text have been corrected, and column {columnToWiden} widened. Saved to {outputFile}.')
        else:
            print('No changes were needed. The file remains unaltered.')

        # Close the workbook
        wb.close()

    except Exception as e:
        print(f'Error: {e}')


def widenColumns(worksheet, columns, width):
    """
    Widen specified columns in an Excel worksheet to ensure data is displayed correctly.

    Args:
        worksheet (openpyxl.worksheet.worksheet.Worksheet): The Excel worksheet to operate on.
        columns (list): A list of column letters (e.g., ['A', 'B', 'C']) to be widened.
        width (int): The desired width for the columns.

    Returns:
        None

    Example:
        To widen columns A, B, and C in a worksheet to a width of 15:
        widenColumns(worksheet, ['A', 'B', 'C'], width=15)
    """
    for col in columns:
        worksheet.column_dimensions[col].width = width
