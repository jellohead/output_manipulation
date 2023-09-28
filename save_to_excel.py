# save_to_excel.py
# open_excel_workbook checks for existing specified workbook and opens it or creates a new
#   workbook if it doesn't exist.
# save_result_to_excel saves a dataframe stored in 'result' and saves it to a specifie excel worksheet
import openpyxl
import pandas as pd


def open_excel_workbook(excel_filename):
    # Open the specified excel workbook, or create one if it doesn't exist
    # variables: excel_filename
    try:
        # Try to open an existing workbook
        workbook = openpyxl.load_workbook(excel_filename)
    except FileNotFoundError:
        # If the file does not exist, create a new workbook
        workbook = openpyxl.Workbook()
        workbook.save(excel_filename)


def save_result_to_xlworksheet(excel_filename, worksheet_name, result):
    # Append the DataFrame to the excel workbook
    # variables: excel_filename, worksheet_name, result
    with pd.ExcelWriter(
        excel_filename, if_sheet_exists="overlay", engine="openpyxl", mode="a"
    ) as writer:
        # result.to_excel(writer, sheet_name=worksheet_name, index_label="Index")
        result.to_excel(
            writer,
            sheet_name=worksheet_name,
            index_label="Index",
            startrow=len(excel_filename) + 1,
        )
