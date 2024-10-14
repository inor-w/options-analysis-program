import logging

from gui import load_gui
from excel_utils import close_workbook, populate_sheets
from data_loader import prepare_data_sheet
from calculations import calc_variables

def main():
    logging.basicConfig(level=logging.INFO, format='%(message)s')
    logging.info("Initializing the program")
    load_gui()
    logging.info("Starting Script")
    close_workbook()
    data_sheet, main_sheet, naked_put_sheet, workbook = prepare_data_sheet()
    calculation_results = calc_variables(data_sheet)
    populate_sheets(calculation_results, data_sheet, main_sheet, naked_put_sheet, workbook)

if __name__ == "__main__":
    main()
