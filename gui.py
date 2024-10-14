import os
import sys
import time
import tkinter as tk
from tkinter import messagebox
import logging

def load_gui():
    """
    Display an initial GUI window to check if the CSV file has been updated.
    """
    csv_file_path = "C:/Users/Inorw/Downloads/Data.csv"
    current_time = time.time()
    try:
        file_modified_time = os.path.getmtime(csv_file_path)
    except FileNotFoundError:
        logging.error(f"The file {csv_file_path} was NOT found")
        messagebox.showerror("File NOT Found", "The Data.csv could NOT be found in the Downloads folder")
        sys.exit()
    except Exception as e:
        logging.error(f"An unexpected error occurred: {str(e)}")
        messagebox.showerror("ERROR", f"An unexpected error occurred: {str(e)}")
        sys.exit()
    else:
        if (current_time - file_modified_time) <= 43200:
            logging.info("SUCCESS. File was modified within 12 hours (43,200 seconds)")
            root = tk.Tk()
            root.withdraw()
            root.title("Options Excel Program - specified to ATP")
            response = messagebox.askyesno("CSV Update Check", "Have you updated the Data.csv file in the Downloads folder?")
            root.destroy()
            if not response:
                messagebox.showinfo("Exiting", "Please update the Data.csv file before running the program. Exiting... - Inor Wang")
                logging.error("User chose no. Stopping script")
                sys.exit()
            else:
                logging.info("SUCCESS. User chose yes. Starting script")
        else:
            logging.error("File was not modified within 12 hours (43,200 seconds)")
            messagebox.showinfo("ERROR", "Data.csv file was NOT modified within 12 hours!")
            sys.exit()
