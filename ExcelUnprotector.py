# Import packages
import os
import zipfile
import shutil
import tempfile
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed

__author__ = "Nguyen Trung Hieu"
__email__ = "trunghieuth10@gmail.com"
__version__ = "1.0"
__date__ = "2024-07-19"
__description__ = """
ExcelUnprotector: This script unlocks protected sheets in Excel files (.xls, .xlsx, .xlsm).
You can specify files or directories via command line arguments, or use a graphical file chooser dialog if no arguments are provided.
"""

# Initialize dictionary to store imported modules
installed_modules = {}

# Function to automatically install and import missing packages
def install_and_import(package):
    import importlib
    import subprocess
    if package not in installed_modules:
        try:
            installed_modules[package] = importlib.import_module(package)
        except ImportError:
            subprocess.check_call(['pip', 'install', package])
            installed_modules[package] = importlib.import_module(package)
    return installed_modules[package]

# Install and import necessary libraries
lxml = install_and_import('lxml')
tqdm = install_and_import('tqdm')
from lxml import etree
from tqdm import tqdm

# Set up logging configuration
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("unlock_excel.log"),
        logging.StreamHandler()
    ]
)

def remove_sheet_protection(file_path, max_workers=10):
    """
    Remove sheet protection in an Excel file.

    :param file_path: Path to the Excel file to unlock.
    :param max_workers: Maximum number of worker threads to use.
    :return: Path to the unlocked Excel file.
    """
    logging.debug(f"Processing file: {file_path}")
    tmp_dir = tempfile.mkdtemp()
    output_file = file_path.replace('.xls', '_unprotected.xls')
    
    try:
        # Check if the Excel file exists
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")

        # Unzip the Excel file to a temporary directory
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(tmp_dir)
        logging.debug(f"Extracted {file_path} to temporary directory {tmp_dir}")
      
        # Find and modify sheet XML files to remove password protection
        worksheets_dir = os.path.join(tmp_dir, 'xl', 'worksheets')
        
        def process_sheet(sheet_file):
            """
            Process a single sheet XML file to remove protection.

            :param sheet_file: Name of the sheet XML file to process.
            """
            sheet_path = os.path.join(worksheets_dir, sheet_file)
            try:
                tree = etree.parse(sheet_path)
                root = tree.getroot()

                # Find and remove <sheetProtection> element
                sheet_protection = root.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheetProtection')
                if sheet_protection is not None:
                    sheet_protection.getparent().remove(sheet_protection)
                    logging.debug(f"Removed sheet protection from {sheet_path}")

                # Save the modified XML file
                tree.write(sheet_path, xml_declaration=True, encoding='UTF-8', pretty_print=True)
            except etree.XMLSyntaxError as e:
                logging.error(f"Error parsing XML file {sheet_path}: {e}")
            except Exception as e:
                logging.error(f"An error occurred while processing sheet {sheet_file}: {e}")
        
        # List all sheet XML files and store in an array
        sheet_files = [f for f in os.listdir(worksheets_dir) if f.endswith(".xml")]  

        # Use progress bar to show the process of handling sheet XML files
        with tqdm(total=len(sheet_files), desc=f"Processing {os.path.basename(output_file)}", dynamic_ncols=True) as pbar:
            # Use multithreading to handle sheet XML files in the array
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = [executor.submit(process_sheet, sheet_file) for sheet_file in sheet_files]
                for future in as_completed(futures):
                    try:
                        future.result()
                    except Exception as e:
                        logging.error(f"An error occurred while processing a sheet: {e}")
                    finally:
                        pbar.update(1)

        # Zip the files back into a new Excel file
        with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
            for root, _, files in os.walk(tmp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, tmp_dir)
                    zip_ref.write(file_path, arcname)        
        logging.debug(f"File unprotected saved as: {output_file}")
        return output_file
        
    except FileNotFoundError as e:
        logging.error(f"File not found: {file_path}")
    except zipfile.BadZipFile as e:
        logging.error(f"Bad Zip file: {file_path}")
    except Exception as e:
        logging.error(f"An error occurred: {e}")
    finally:
        # Delete temporary directory if it exists
        if os.path.exists(tmp_dir):
            shutil.rmtree(tmp_dir)
            logging.debug(f"Deleted temporary directory {tmp_dir}")

def unlock_excel_sheets(input_path, max_workers=None):
    """
    Unlock protected sheets in Excel files or all Excel files in a directory.

    :param input_paths: Path to a file or directory containing Excel files.
    :param max_workers: Maximum number of worker threads to use.
    :return: Message containing paths to the unlocked Excel files.
    """
    excel_files = []
    if os.path.isfile(input_path) and input_path.endswith((".xls", ".xlsx", ".xlsm")):
        # If input is an Excel file
        excel_files.append(input_path)
    elif os.path.isdir(input_path):
        # If input is a directory
        for root, _, files in os.walk(input_path):
            for file in files:
                if file.endswith((".xls", ".xlsx", ".xlsm")):
                    file_path = os.path.join(root, file)
                    excel_files.append(file_path)
    else:
        logging.error(f"Invalid input path: {input_path}")
        return

    list_outfile = ""
    if len(excel_files) > 1:
        pbar = tqdm(total=len(excel_files), desc=f"==>Total process: {len(excel_files)} files", dynamic_ncols=True, position=0)
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(remove_sheet_protection, file_path) for file_path in excel_files]
        for future in as_completed(futures):
            try:
                list_outfile += future.result() + "\n"
            except Exception as e:
                logging.error(f"An error occurred while processing a file: {e}")
            finally:
                if len(excel_files) > 1: pbar.update(1)
    return list_outfile

def filedialog_input():
    """
    Open a file dialog to select Excel files or directories containing Excel files.
    
    :return: List of selected file or directory paths.
    """
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()

    file_paths = filedialog.askopenfilenames(title="Select excel files to unprotect - trunghieuth10", filetypes=[("Excel files", "*.xls;*.xlsx;*.xlsm")])
    if not file_paths:
        folder_path = filedialog.askdirectory(title="Select folder contains excel files to unprotect - trunghieuth10")
        if folder_path:
            file_paths = [folder_path]
        else:
            logging.error("No file or directory selected")
            return
        
    return file_paths

def parse_input():
    """
    Parse command line arguments or open a file dialog if no arguments are provided.
    
    :return: List of input paths.
    """
    import argparse
    parser = argparse.ArgumentParser(
        description='Unlock protected sheets in Excel files.',
        epilog='Example usage:\n'
               '  python unlock_excel.py path/to/your/excel/file.xlsx\n'
               '  python unlock_excel.py path/to/your/excel/folder\n'
               '  python unlock_excel.py path/to/your/excel/file.xlsx path/to/your/excel/folder',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument(
        'input_path',
        nargs='*',
        help='Path to files or directory containing Excel files. If a directory is provided, all Excel files in the directory and its subdirectories will be processed.'
    )
    args = parser.parse_args()

    if args.input_path:
        return args.input_path
    
    return filedialog_input()

def main():
    input_paths = parse_input()
    if not input_paths:
        exit()
    for input_path in input_paths:
        if os.path.exists(input_path):
            result = unlock_excel_sheets(input_path, max_workers=None)
            # if result:
            #     print(f"File unprotected saved as:\n{result}")
        else:
            logging.error(f"Input path does not exist: {input_path}")

if __name__ == "__main__":
    main()
