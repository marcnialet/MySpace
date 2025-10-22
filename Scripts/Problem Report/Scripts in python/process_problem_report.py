import logging
import argparse
import os
import csv
import zipfile
import re
import math
import shutil
import sys
import fnmatch
import msvcrt
import glob
import xml.etree.ElementTree as ET
from datetime import datetime

def find_dates_in_files(path, pattern):
    # Initialize variables to store the overall oldest and youngest dates
    overall_oldest_date = None
    overall_youngest_date = None

    # Create the search path
    search_path = os.path.join(path, pattern)
    
    # Get the list of files matching the pattern
    files = glob.glob(search_path)
    
    # Check if there are any files to process
    if not files:
        print("No files found matching the pattern.")
        return

    # Total number of files for progress tracking
    total_files = len(files)

    # Iterate over each file matching the pattern with a progress indicator
    for i, file in enumerate(files):
        file_name = os.path.basename(file)

        # Initialize variables to store the oldest and youngest dates for the current file
        file_oldest_date = None
        file_youngest_date = None

        with open(file, 'r') as f:
            for line in f:
                # Extract the date from the line
                date_str = line.split(',')[0]
                try:
                    date = datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')
                except ValueError:
                    continue  # Skip lines that don't match the date format

                # Update file oldest and youngest dates
                if file_oldest_date is None or date < file_oldest_date:
                    file_oldest_date = date
                if file_youngest_date is None or date > file_youngest_date:
                    file_youngest_date = date

        # Calculate the difference for the current file
        if file_oldest_date and file_youngest_date:
            file_time_diff = file_youngest_date - file_oldest_date

            # Format the difference as Days:Hours:Minutes:Seconds
            total_seconds = int(file_time_diff.total_seconds())
            days, remainder = divmod(total_seconds, 86400)
            hours, remainder = divmod(remainder, 3600)
            minutes, seconds = divmod(remainder, 60)
            formatted_file_diff = f"{days}d:{hours}h:{minutes}m:{seconds}s"
        else:
            formatted_file_diff = "N/A"

        # Update overall oldest and youngest dates
        if file_oldest_date and (overall_oldest_date is None or file_oldest_date < overall_oldest_date):
            overall_oldest_date = file_oldest_date
        if file_youngest_date and (overall_youngest_date is None or file_youngest_date > overall_youngest_date):
            overall_youngest_date = file_youngest_date

        # Combined progress trace for the current file
        print(f"File {i + 1} of {total_files}: {file_name}\t Start: {file_oldest_date} End: {file_youngest_date} Difference: {formatted_file_diff}")

    # Check if we have found any dates overall
    if overall_oldest_date and overall_youngest_date:
        # Calculate the difference
        time_diff = overall_youngest_date - overall_oldest_date

        # Format the difference as Days:Hours:Minutes:Seconds
        total_seconds = int(time_diff.total_seconds())
        days, remainder = divmod(total_seconds, 86400)
        hours, remainder = divmod(remainder, 3600)
        minutes, seconds = divmod(remainder, 60)
        formatted_diff = f"{days}d:{hours}h:{minutes}m:{seconds}s"

        # Print the overall results
        print(f"Overall oldest date: {overall_oldest_date}")
        print(f"Overall youngest date: {overall_youngest_date}")
        print(f"Overall difference: {formatted_diff}")
    else:
        print("No valid dates found in the specified files.")
        
def reformat_backendlogs_linefeeds(folder_path, pattern):
    # Get the list of all files matching the pattern in the specified folder
    files = [f for f in os.listdir(folder_path) if fnmatch.fnmatch(f, pattern)]
    total_files = len(files)

    # Loop through all matching files in the folder
    for index, filename in enumerate(files, start=1):
        file_path = os.path.join(folder_path, filename)
        
        # Display progress at the beginning
        print(f"Reformatting linefeeds for backend log file: {index}/{total_files}: {filename}")

        # Read the content of the file
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()

        # Replace the special character \r by an empty char
        content = content.replace('\r', '')
        content = content.replace('\n', '')

        # Save the modified content back to the file
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(content)

        # Reload the file to perform the next replacement
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()

        # Replace the string [EOL] by [EOL]\r\n
        content = content.replace('[EOL]', '[EOL]\r')

        # Save the modified content back to the file
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(content)

def CountWordInFile(filepath, word, casesensitive, exactword):
    try:
        count = 0
        with open(filepath, 'r') as file:
            for line in file:
                if not casesensitive:
                    line = line.lower()
                words = line.split()
                for w in words:
                    if not casesensitive:
                        w = w.lower()
                    if exactword:
                        if w == word:
                            count += 1
                    else:
                        if word in w:
                            count += 1
        return count
    except (FileNotFoundError, IsADirectoryError, PermissionError):
        return -1
    except UnicodeDecodeError:
        return -1


def create_list_of_files(startpath):
    
    file_path = 'problemreport_files.csv'
    if os.path.exists(file_path):
        os.remove(file_path)
        
    with open(file_path, 'w', newline='') as csvfile:
            filewriter = csv.writer(csvfile, delimiter=';', quotechar='|', quoting=csv.QUOTE_MINIMAL)
            filewriter.writerow(['Name', 'Type', 'File ID', 'Index', '#Errors', '#Exceptions', 'Folder'])

            index = 0            
            for root, dirs, files in os.walk(startpath):
                if root == startpath:  # Include the root folder at the top
                    filewriter.writerow([os.path.basename(root), 'Folder', '', index, '', '', os.path.relpath(root, startpath)])
                    index += 1
                else:
                    filewriter.writerow([os.path.basename(root), 'Folder', '', index, '', '', os.path.relpath(root, startpath)])
                    index += 1
                for file in files:
                    filewriter.writerow([file, 'File', '', index, '', '', os.path.relpath(root, startpath)])
                    index += 1

# Unzips a file with password. Path should be relative
def unzip_file(zip_path, password):
    # Get the absolute path of the zip file
    abs_zip_path = os.path.abspath(zip_path)
    
    # Get the directory where the zip file is located
    unzip_dir = os.path.dirname(abs_zip_path)
    
    # Print the absolute path of the zip file and the output folder path
    print("Zip file path:", abs_zip_path)
    print("Output folder path:", unzip_dir)
    
    # Extract the zip file to the unzip directory
    with zipfile.ZipFile(abs_zip_path, 'r') as zip_ref:
        # Get the total number of files in the zip
        total_files = len(zip_ref.infolist())
        extracted_files = 0
        
        for file_info in zip_ref.infolist():
            # Update progress
            extracted_files += 1
            progress = (extracted_files / total_files) * 100
            print(f"Progress: {progress:.2f}%")
            
            # Extract the file
            file_path = os.path.join(unzip_dir, file_info.filename)
            if os.path.exists(file_path):
                os.remove(file_path)  # Remove existing file before extraction
            if password:
                zip_ref.extract(file_info, unzip_dir, pwd=password.encode())
            else:
                zip_ref.extract(file_info, unzip_dir)
            
            # Print the destination file path
            print(f"Destination File: {file_path}")

# Splits all the LabCoag C4C backend log files in files that are smaller than 20.000 KB, files are saved in a PART folder
def split_log_files(folder_path,partfoldername,pattern,basepart):
    # Create the parts subfolder
    parts_folder = os.path.join(folder_path,partfoldername)
    parts_folder_fullpath = os.path.abspath(parts_folder)
    
    if os.path.exists(parts_folder_fullpath):
        shutil.rmtree(parts_folder_fullpath)
    os.makedirs(parts_folder_fullpath)
    
    # Get the list of log files matching the pattern
    
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
                part_filename = f"{basepart}.{part_number}.part.{str(i+1).zfill(3)}"
                part_filepath = os.path.join(parts_folder_fullpath, part_filename)
                
                with open(part_filepath, 'w') as part_file:
                    part_file.write(part)
                
                print(f"Progress: {i+1}/{part_count} - {part_filepath} (Source File: {log_file_path})")
    
    print("Splitting log files completed.")

# opens an file (ususally an XML file) and overwrites its in UTF-8 encoding
def save_xml_with_utf8_encoding(file_path):
    # Load the XML file
    tree = ET.parse(file_path)
    
    # Get the root element of the XML
    root = tree.getroot()
    
    # Save the XML file with UTF-8 encoding (overwrite the current file)
    tree.write(file_path, encoding="utf-8", xml_declaration=True)
    
    print(f"XML file saved with UTF-8 encoding: {file_path}")

 # finds any ZIP file contained in the input folder
def find_zip_file(folder_path):
    # Get the absolute path of the folder
    abs_folder_path = os.path.abspath(folder_path)
    
    # Search for zip files under the folder
    for root, dirs, files in os.walk(abs_folder_path):
        for file in files:
            if file.endswith(".zip"):
                # Return the path of the zip file appended to the folder path
                return os.path.join(root, file)
    
    # If no zip file is found, return None
    return None

# The problem report file name  can be passed by argument or it can be detected automatically in the root folder of this script

# Create the argument parser
parser = argparse.ArgumentParser(description='process_problem_report - unzips the problem report, unzips the labcore logging report, unzips the ICLogs and prepare the data to be loaded in the GSheet.')

# Add the parameters
parser.add_argument('zipfile', type=str, nargs='?', const=None, default=None, help='Relative Path to the zip file. If empty it looks for a zip file on the same folder')
parser.add_argument('-all', action='store_true', default=True, help='Unzips all zips files in the problem report, the PR itself and also the LabCore ZIP and also IC zip')
parser.add_argument('-nounzip', action='store_true', default=False, help='Do not unzips the problem report it self')
parser.add_argument('-nounzip_loggins', action='store_true', default=False, help='Do not unzip the the LabCore logging zip')
parser.add_argument('-unzip_ics', action='store_true', default=False, help='Do unzips the IC Log')
parser.add_argument('-nounzip_settings', action='store_true', default=False, help='Do not unzip the the Settings file zip')

# Set the usage description
parser.usage = 'python process_problem_report.py <path_to_zidddpfile> [-all] [-nounzip] [-nounzip_loggins] [-unzip_ics]'

# Parse the arguments
args = parser.parse_args()

# Access the path of the zip file
zipfile_path = args.zipfile

# Map the other parameters to variables
all_param = args.all
nounzip_param = args.nounzip
nounzip_loggins_param = args.nounzip_loggins
unzip_ics_param = args.unzip_ics
nounzip_settings_param = args.nounzip_settings

# Print the mapped values
print(f"Zipfile path: {zipfile_path}")
print(f"All parameter: {all_param}")
print(f"No Unzip parameter: {nounzip_param}")
print(f"No Unzip loggings parameter: {nounzip_loggins_param}")
print(f"No Unzip ICS parameter: {unzip_ics_param}")
print(f"No Unzip Settings parameter: {nounzip_settings_param}")

passwordPR = "c0m1t@Pr0bl3mR3p0rt"  
passwordSettings = "c4c@3xp0r7And1mp0r7"


###### STEP 1 UNZIP
# Unzip the problem report

# Check if the zip file path is provided as a command-line argument  (optional by parameter, default true)
if zipfile_path is None or not os.path.exists(zipfile_path):
    problemreport_zip_file = find_zip_file(r".\/")
else: 
    problemreport_zip_file = zipfile_path
    
if not os.path.exists(problemreport_zip_file):
    print("Please provide the problem report zip file name, or exectute this script into a folder containing  a problem report zip file")
    sys.exit(1)


if all_param and not nounzip_param:
    unzip_file(problemreport_zip_file, passwordPR)
else: print("unzip problem report skipped")

###### STEP 2 Reformat line feeds: removes all line feeds then add a line feed after [EOL] -> each log one line
# Specify the folder path and pattern here
folder_path = r"./EventTraces/SystemSoftwareLog"
pattern = 'Roche.C4C.ServiceHosting.BackEnd*'
# removes all carriage return, then after the symbol [EOL] adds a carriage return.
reformat_backendlogs_linefeeds(folder_path, pattern)
###### STEP 3: Traces the duration in time of each log file and the global duration 
# Looks for the yougest oldest date in the log file and the duration as well as the global duration
find_dates_in_files(folder_path, pattern)

###### STEP 4 SPLIT LOG FILES into smaller parts in another folder 
# Split backend log files in smaller files

folder_path = r".\EventTraces\SystemSoftwareLog"
pattern = r"Roche\.C4C\.ServiceHosting\.BackEnd\.log(\.\d{0,3})?$"
basepart = r"Roche.C4C.ServiceHosting.BackEnd.log"
partfoldername = "parts_backendlog"
# Call the split_log_files function to split each backend log file in file max 20.000KB
split_log_files(folder_path,partfoldername,pattern, basepart)

folder_path = r".\EventTraces\SystemSoftwareLog\ComitServices"
pattern = r"ServiceHosting\.ComitServices\.log(\.\d{0,3})?$"
partfoldername = "parts_ServiceHosting"
basepart = r"ServiceHosting.ComitServices.log"
# Call the split_log_files function to split each backend log file in file max 20.000KB
split_log_files(folder_path,partfoldername,pattern, basepart)

folder_path = r".\EventTraces\SystemSoftwareLog\ComitServices"
pattern = r"SystemServiceHost\.ComitServices\.log(\.\d{0,3})?$"
partfoldername = "parts_SystemServiceHost"
basepart = r"SystemServiceHost.ComitServices.log"
# Call the split_log_files function to split each backend log file in file max 20.000KB
split_log_files(folder_path,partfoldername,pattern, basepart)

###### STEP 5 Unzip audits and user messages  (optional by parameter, default true)
# Unzip the zip files that contains the LabCore Audits and Messages (relative folder is always the same)
folder_path = r".\EventTraces\Logging" 
# Call the find_zip_file function looks for the zip files that contains the audits and messages
logging_zip_file_path = find_zip_file(folder_path)
if logging_zip_file_path:
    print("Audits and messages settings: zip file found:", logging_zip_file_path)
    if all_param and not nounzip_loggins_param:
        unzip_file(logging_zip_file_path, "") #unzip the fils that contains the audits and messages
    else: print("unzip logging_zip_file_path skipped")
else:
    print("No zip file found in the folder. ", logging_zip_file_path)

###### STEP 6 Fix audit and messagges XML file to ensure it has UTF-8 Encoding
# Provide the relative path of the XML file as the parameter
logging_file_path = r".\EventTraces\Logging\LoggingExport.xml"
if logging_file_path:
    # Call the save_xml_with_utf8_encoding function
    save_xml_with_utf8_encoding(logging_file_path)    
else:
    print("File not found: ", logging_file_path)

###### STEP 7 Unzip the IC Log (optional by parameter, default false)
ic_logs_file_path = r".\EventTraces\IC\log\log.zip"
if all_param and unzip_ics_param:
    unzip_file(ic_logs_file_path, "") #unzip the fils that contains the audits and messages
else: print("unzip iclogs skipped")

###### STEP 8 Unzip settings (optional by parameter, default true)
# Unzip the zip files that contains the application exported settings (relative folder is always the same)
folder_path = r".\EventTraces\SystemConfigurationExport" 
# Call the find_zip_file function looks for the zip files that contains the settings
logging_zip_file_path = find_zip_file(folder_path)
if logging_zip_file_path:
    print("Application settings: zip file found:", logging_zip_file_path)
    if all_param and not nounzip_settings_param:
        unzip_file(logging_zip_file_path, passwordSettings) #unzip the fils that contains the audits and messages
    else: print("unzip logging_zip_file_path skipped")
else:
    print("No zip file found in the folder. ", logging_zip_file_path)

###### STEP 9 Create cvs files with all files in the PR
create_list_of_files('.')

print("Press any key to continue")
msvcrt.getch()