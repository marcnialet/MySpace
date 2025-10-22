
import os
import re
import math
import shutil

def split_log_files(folder_path):
    # Create the parts subfolder
    parts_folder = os.path.join(folder_path, "parts")
    parts_folder_fullpath = os.path.abspath(parts_folder)
    
    if os.path.exists(parts_folder_fullpath):
        shutil.rmtree(parts_folder_fullpath)
    os.makedirs(parts_folder_fullpath)
    
    # Get the list of log files matching the pattern
    pattern = r"Roche\.C4C\.ServiceHosting\.BackEnd\.log(\.\d{0,3})?$"
    log_files = [file for file in os.listdir(folder_path) if re.match(pattern, file)]
    
    # Process each log file
    for log_file in log_files:
        log_file_path = os.path.join(folder_path, log_file)
        part_count = 1
        
        with open(log_file_path, 'r') as file:
            content = file.read()
            file_size = os.path.getsize(log_file_path)
            max_part_size = 20000 * 1024  # Maximum size in bytes
            
            # Calculate the number of parts needed
            part_count = math.ceil(file_size / max_part_size)
            
            # Split the content into parts
            parts = []
            remaining_content = content
            
            for i in range(part_count):
                part_size = min(len(remaining_content), max_part_size)
                part = remaining_content[:part_size]
                
                # Check if the part ends with a carriage return
                if i < part_count - 1:
                    last_newline_index = part.rfind("\n")
                    if last_newline_index != -1:
                        part = part[:last_newline_index + 1]
                        remaining_content = remaining_content[last_newline_index + 1:]
                
                parts.append(part)
            
            # Write each part to a new file
            for i, part in enumerate(parts):
                part_number = log_file.split(".")[-1].zfill(3) if log_file.split(".")[-1].isdigit() else "000"
                part_filename = f"Roche.C4C.ServiceHosting.BackEnd.log.{part_number}.part.{str(i+1).zfill(3)}"
                part_filepath = os.path.join(parts_folder_fullpath, part_filename)
                
                with open(part_filepath, 'w') as part_file:
                    part_file.write(part)
                
                print(f"Progress: {i+1}/{part_count} - {part_filepath} (Source File: {log_file_path})")
    
    print("Splitting log files completed.")

# Provide the relative path of the folder containing the log files as the parameter
folder_path = r".\EventTraces\SystemSoftwareLog"

# Call the split_log_files function
split_log_files(folder_path)
