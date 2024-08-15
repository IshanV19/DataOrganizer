import os
import pandas as pd
from datetime import datetime
import re

def extract_positive_time_unit(sample):
    """Extracts positive time units from the sample string."""
    match = re.search(r'([dwm])(\d+)', sample) 
    if match:
        unit = match.group(1)
        value = int(match.group(2))

        if unit == 'd':
            return f"{value}"
        elif unit == 'w':
            return f"{value * 7}" 
        elif unit == 'm':
            return f"{value * 30}"
    return None

def extract_negative_time_unit(sample):
    """Extracts negative time units from the sample string."""
    match = re.search(r'([dwm])-?(\d+)', sample) 
    if match:
        unit = match.group(1)
        value = int(match.group(2))

        if unit == 'd':
            return f"{-value}"
        elif unit == 'w':
            return f"{-value * 7}"
        elif unit == 'm':
            return f"{-value * 30}"
    return None

def extract_time_unit(sample):
    """Extracts time units from the sample string, handling both positive and negative cases."""
    if re.search(r'[dwm]-\d+', sample):
        return extract_negative_time_unit(sample)
    elif re.search(r'[dwm]\d+', sample):
        return extract_positive_time_unit(sample)
    return None

def trim_sample_name(sample):
    """Trim parts of the sample name that start with 'm', 'w', 'd', or match '1-2'."""
    if isinstance(sample, str):
        parts = sample.split()
        trimmed_parts = [part for part in parts if not (part.startswith('m') or part.startswith('w') or part.startswith('d') or part == '1-2')]
        return ' '.join(trimmed_parts)
    return sample


def determine_sample_group(sample_series, base_id=101, increment=100):
    """Determines the sample group based on the 'Sample' column content."""
    sample_groups = {}
    
    for sample in sample_series:
        if isinstance(sample, str):
            match = re.match(r'(\d+)', sample)
            if match:
                id_prefix = int(match.group(1))
                if id_prefix >= base_id:
                    group_id = ((id_prefix - base_id) // increment) * increment + base_id
                    if group_id not in sample_groups:
                        sample_groups[group_id] = 0
                    sample_groups[group_id] += 1
    
    if sample_groups:
        dominant_group = max(sample_groups, key=sample_groups.get)
        return f"{dominant_group}"
    return None

def combine_and_organize(input_directory, output_directory, processed_files=None):
    
    if processed_files is None:
        processed_files = []

    try:
        # Make sure output directory exists
        if not os.path.exists(output_directory):
            os.makedirs(output_directory)
        
        # Get all organized data files
        organized_files = [f for f in os.listdir(input_directory) if f.endswith('_organized_data.xlsx') and f not in processed_files]
        
        print(f"Found organized files: {organized_files}")  # Debugging line
        
        if not organized_files:
            print(f"No new organized data files found in '{input_directory}'. Exiting.")
            return
        
        master_dfs = {}

        for file in organized_files:
            file_path = os.path.join(input_directory, file)
            df = pd.read_excel(file_path, sheet_name=None)  

            for sheet_name, sheet_df in df.items():
                if 'Sample' in sheet_df.columns:
                    sheet_df['Days'] = sheet_df['Sample'].apply(extract_time_unit)
                    sheet_df['Sample'] = sheet_df['Sample'].apply(trim_sample_name)

                    sample_group = determine_sample_group(sheet_df['Sample'])

                    if sample_group:
                        if sample_group not in master_dfs:
                            master_dfs[sample_group] = pd.DataFrame()
                        
                        sheet_df = sheet_df.copy()  
                        sheet_df['Data Source'] = file  
                        sheet_df['Time Added'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # Timestamp of when added

                        master_dfs[sample_group] = pd.concat([master_dfs[sample_group], sheet_df], ignore_index=True)

            # Mark this file as processed
            processed_files.append(file)

        # Remove specified columns and save each Master Table to its own Excel file
        columns_to_remove = ['Recovery_1', 'Recovery_2', 'Avg_Recovery', 'Std_Dev_Recovery']
        for sample_group, master_df in master_dfs.items():
            if not master_df.empty:
                # Remove rows starting with C0 or S0 in 'Sample' column
                master_df = master_df[~master_df['Sample'].astype(str).str.startswith(('C0', 'S0'))].copy()

                # Remove rows where 'Sample' column contains 'neat' (case insensitive)
                master_df = master_df[~master_df['Sample'].str.contains('neat', case=False, na=False)].copy()

                # Remove specified columns
                master_df.drop(columns=columns_to_remove, errors='ignore', inplace=True)

                # Fill any blank cells with 0
                master_df.fillna(0, inplace=True)

                # Filter data for the current group 
                filtered_df = master_df[master_df['Sample'].astype(str).str.startswith(sample_group)]

                # Save to Excel with specified sheet name
                master_table_file = os.path.join(output_directory, f'{sample_group}_master_table.xlsx')
                with pd.ExcelWriter(master_table_file) as writer:
                    filtered_df.to_excel(writer, sheet_name=f'{sample_group}_master', index=False)
                print(f"Master table for sample group '{sample_group}' updated successfully at '{master_table_file}'.")

    except Exception as e:
        print(f"An error occurred while combining data into master tables: {e}")
