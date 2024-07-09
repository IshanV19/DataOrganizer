# FileHandler.py

import os

def ask_directory(description):
    """Prompt user to input a directory path."""
    return input(f"{description} ").strip()

def check_directory(directory):
    """Check if the directory exists and is valid."""
    if not os.path.exists(directory):
        raise FileNotFoundError(f"Directory '{directory}' does not exist.")
    if not os.path.isdir(directory):
        raise NotADirectoryError(f"'{directory}' is not a directory.")

def select_excel_files(directory):
    """Select .xlsx or .csv files to process."""
    valid_extensions = ['.xlsx', '.csv']  
    files = [f for f in os.listdir(directory) if os.path.splitext(f)[1] in valid_extensions]
    print("Select the .xlsx or .csv files to process (comma-separated, enter '0' to cancel):")
    for i, file in enumerate(files, start=1):
        print(f"{i}. {file}")

    selections = input("Enter file numbers (e.g., 1, 2, 3) or 'a' for all: ").strip()
    if selections == '0':
        return []

    if selections.lower() == 'a':
        return files

    selected_files = []
    try:
        selected_indices = [int(idx.strip()) - 1 for idx in selections.split(',')]
        selected_files = [files[idx] for idx in selected_indices]
    except (ValueError, IndexError):
        print("Invalid selection. Processing all .xlsx and .csv files.")

    return selected_files

def confirm_action():
    """Prompt user to confirm an action (e.g., adding data to master table)."""
    while True:
        response = input("Would you like to add the organized data to the master table? (yes/no): ").strip().lower()
        if response in ['yes', 'y']:
            return True
        elif response in ['no', 'n']:
            return False
        else:
            print("Invalid input. Please enter 'yes' or 'no'.")
