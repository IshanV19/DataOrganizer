import os
import csv
from datetime import datetime

def ask_directory(description):
    """Prompt user to input a directory path."""
    return input(f"{description} ").strip()

def check_directory(directory):
    """Check if the directory exists and is valid."""
    if not os.path.exists(directory):
        raise FileNotFoundError(f"Directory '{directory}' does not exist.")
    if not os.path.isdir(directory):
        raise NotADirectoryError(f"'{directory}' is not a directory.")

def get_all_csv_files(directory):
    """Get all .csv files in the directory."""
    valid_extensions = ['.csv']
    return [f for f in os.listdir(directory) if os.path.splitext(f)[1].lower() in valid_extensions]

def update_tracker_file(tracker_file, file_name, processed_date, last_modified_date):
    """Update the tracker file with the processed date and last modified date."""
    if os.path.exists(tracker_file):
        # Read existing data
        with open(tracker_file, 'r', newline='') as file:
            reader = list(csv.reader(file))
    else:
        reader = [['File Name', 'File Processed', 'File Update']]
    
    # Update or append new entry
    updated = False
    for row in reader:
        if row[0] == file_name:
            row[1] = processed_date
            row[2] = last_modified_date
            updated = True
            break

    if not updated:
        reader.append([file_name, processed_date, last_modified_date])
    
    # Write back to the tracker file
    with open(tracker_file, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerows(reader)

def get_last_modified_date(file_path):
    """Get the last modified date of a file."""
    return datetime.fromtimestamp(os.path.getmtime(file_path)).strftime("%Y-%m-%d %H:%M:%S")

def read_paths_from_file(file_path):
    """Read input and output paths from the specified file."""
    with open(file_path, 'r') as file:
        lines = file.readlines()

    input_address = ''
    output_address_organized = ''
    output_address_master = ''
    for line in lines:
        if line.startswith('input address: '):
            input_address = line.split(':', 1)[1].strip()
        elif line.startswith('output address organized: '):
            output_address_organized = line.split(':', 1)[1].strip()
        elif line.startswith('output address master: '):
            output_address_master = line.split(':', 1)[1].strip()

    return input_address, output_address_organized, output_address_master