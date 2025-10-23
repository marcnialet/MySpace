import sys
import os
import re
import zipfile
import csv
import fnmatch
import glob
import math
import shutil
import subprocess
from datetime import datetime
import time

# Global variable for the output folder name
OUTPUT_FOLDER_NAME = "Process output"  # Base folder where all output CSV files will be stored

# Global variables for all output CSV file names
PROCESS_LOG_FILE            = "process_log_output.csv"          # CSV that holds all log messages
PROBLEM_REPORT_FILE         = "problemreport_files.csv"         # CSV listing all files in the problem report directory
PIPETTING_EVENTS_FILE       = "testorders_pipetting_events.csv" # CSV storing extracted pipetting events
EBARCODE_SUMMARY_FILE       = "eBarcode_Importer_Summary.csv"   # CSV summarizing e-Barcodes importer analysis
IC_LOG_SUMMARY_FILE         = "IC_Log_Summary.csv"              # CSV summarizing IC log extraction

def setup_output_folder(root_folder):
    """
    Ensures the output folder exists in the root folder.
    Returns the full path to this folder.
    """
    output_folder = os.path.join(root_folder, OUTPUT_FOLDER_NAME)
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        log_message(f"Created output folder: {output_folder}")
    return output_folder
    

def log_message(message):
    """
    Logs a message with a timestamp. Appends it to the log CSV file in the output folder.
    Skips writing if the file is locked (e.g., the log file is open in Excel).
    """
    # Common variables
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] {message}"
    print(log_entry)  # Print to console

    # Path of the global log file
    output_folder = setup_output_folder(os.getcwd())
    file_path = os.path.join(output_folder, PROCESS_LOG_FILE)
    file_exists = os.path.exists(file_path)

    max_retries = 3  # Number of retries in case of file lock
    retry_delay = 10  # Delay in seconds between retries

    for attempt in range(max_retries):
        try:
            with open(file_path, mode="a", newline="", encoding="utf-8-sig") as log_file:
                log_writer = csv.writer(log_file, delimiter=';')  # Use `;` for Excel compatibility
                if not file_exists:
                    log_writer.writerow(["Date", "Message"])  # Add header if it's a new file
                    file_exists = True
                log_writer.writerow([timestamp, message])
            break  # Exit loop if successful
        except PermissionError:
            # Retry writing if the file is locked
            if attempt + 1 < max_retries:
                print(f"Could not write to '{file_path}' after attempt {attempt + 1}, waiting...")
                time.sleep(retry_delay)  # Wait before retrying
            else:
                print(f"Could not write to '{file_path}' after {max_retries} retries.")
                break
        except Exception as e:
            print(f"An unexpected error occurred while trying to log to '{file_path}': {e}")
            break
            
def extract_problem_report(root_folder):
    """
    Processes the problem report:
    1. Unzips the problem report.
    2. Reformats linefeeds in log files.
    3. Finds and logs the oldest and youngest dates in log files.
    """
    ###### STEP 1: Unzip the problem report ######
    log_message("Starting problem report extraction process.")
    passwordPR = "c0m1t@Pr0bl3mR3p0rt"
    problemreport_zip_file = find_zip_file(root_folder)

    if not problemreport_zip_file or not os.path.exists(problemreport_zip_file):
        log_message("No zip file found for the problem report. Ensure the file exists in the root folder.")
        return

    log_message(f"Unzipping problem report: {problemreport_zip_file}")
    unzip_file(problemreport_zip_file, passwordPR)

    ###### STEP 2: Reformat linefeeds ######
    log_message("Reformatting line feeds in backend log files...")
    folder_path = os.path.join(root_folder, "EventTraces\\SystemSoftwareLog")
    pattern = "Roche.C4C.ServiceHosting.BackEnd*"
    reformat_backendlogs_linefeeds(folder_path, pattern)
    log_message("Line feed reformatting completed.")

    ###### STEP 3: Find and log dates ######
    log_message("Tracing the duration in time for each log file and calculating global duration...")
    find_dates_in_files(folder_path, pattern)
    log_message("Date tracing and duration calculation completed.")

    log_message("Problem report extraction process completed.")

def split_log_files_in_parts(root_folder):
    """
    Splits backend log files into smaller parts (max size: 20,000 KB each) and saves them
    into respective "parts" folders for each type of log file.
    """
    log_message("Starting process to split backend log files into smaller parts.")
    
    # List of folders, part folder names, and patterns
    folder_paths = [
        (os.path.join(root_folder, ".\\EventTraces\\SystemSoftwareLog"), "parts_backendlog", 
         r"^Roch\e\.C4C\.ServiceHosting\\.BackEnd\.log.*$", "Roche.C4C.ServiceHosting.BackEnd.log"),
        
        (os.path.join(root_folder, ".\\EventTraces\\SystemSoftwareLog\\ComitServices"), "parts_ServiceHosting", 
         r"^ServiceHosting\\.ComitServices\\.log.*$", "ServiceHosting.ComitServices.log"),
        
        (os.path.join(root_folder, ".\\EventTraces\\SystemSoftwareLog\\ComitServices"), "parts_SystemServiceHost", 
         r"^SystemServiceHost\\.ComitServices\\.log.*$", "SystemServiceHost.ComitServices.log")
    ]

    for folder_path, partfoldername, pattern, basepart in folder_paths:
        log_message(f"Processing folder: {os.path.abspath(folder_path)}")
        # Compile regex for pattern matching
        regex_pattern = re.compile(pattern)

        # Call split_log_files with adjusted parameters
        split_log_files(folder_path, partfoldername, regex_pattern, basepart)

    log_message("Completed splitting backend log files into smaller parts.")

def extract_audits(root_folder):
    """
    Extracts audits and user messages:
    1. Unzips the audits and user messages zip file.
    2. Fixes the extracted audit and message XML file to ensure UTF-8 encoding.
    """
    log_message("Starting process to extract audits and user messages.")

    # STEP 5: Unzip audits and user messages
    folder_path = os.path.join(root_folder, "EventTraces\\Logging")
    log_message(f"Looking for zip file in: {folder_path}")
    logging_zip_file_path = find_zip_file(folder_path)

    if logging_zip_file_path:
        log_message(f"Found audits and messages zip file: {logging_zip_file_path}")
        unzip_file(logging_zip_file_path, "")
        log_message("Unzipping audits and user messages completed.")
    else:
        log_message("No zip file found for audits and user messages.")
        return  # Skip if no zip file is found

    # STEP 6: Fix audit and message XML file encoding
    logging_file_path = os.path.join(folder_path, "LoggingExport.xml")
    log_message(f"Looking for audit messages XML file: {logging_file_path}")

    if os.path.exists(logging_file_path):
        save_xml_with_utf8_encoding(logging_file_path)
        log_message("XML file encoding fixed to UTF-8.")
    else:
        logging_file_path = os.path.join(folder_path, "logEventsData.xml")
        log_message(f"Looking for audit messages XML file: {logging_file_path}")
        if os.path.exists(logging_file_path):
            save_xml_with_utf8_encoding(logging_file_path)
            log_message("XML file encoding fixed to UTF-8.")   
        else:
            log_message("Audit messages XML file not found. Skipping UTF-8 encoding fix.")

    log_message("Audit and user messages extraction process completed.")

def extract_settings(root_folder):
    """
    Extracts application settings:
    1. Unzips the settings zip file located in the SystemConfigurationExport folder.
    """    
    log_message("Starting process to extract application settings.")

    # STEP 8: Unzip settings
    folder_path = os.path.join(root_folder, "EventTraces\\SystemConfigurationExport")
    log_message(f"Looking for zip file in: {folder_path}")
    logging_zip_file_path = find_zip_file(folder_path)

    if logging_zip_file_path:
        log_message(f"Found settings zip file: {logging_zip_file_path}")
        passwordSettings = "c4c@3xp0r7And1mp0r7"
        unzip_file(logging_zip_file_path, passwordSettings)
        log_message("Unzipping application settings completed.")
    else:
        log_message("No settings zip file found in the folder. Skipping extraction.")

    log_message("Application settings extraction process completed.")

def extract_ic_log(root_folder):
    """
    Extracts the IC log into the 'Process output/' directory.
    """    
    log_message("Starting process to extract IC logs.")
    output_folder = setup_output_folder(root_folder)

    ic_logs_file_path = os.path.join(root_folder, "EventTraces\\IC\\log\\log.zip")
    log_message(f"Looking for IC log zip file at: {ic_logs_file_path}")

    # Check if the IC log exists
    if os.path.exists(ic_logs_file_path):
        log_message(f"Found IC log zip file: {ic_logs_file_path}")
        unzip_file(ic_logs_file_path, "")
        csv_summary_path = os.path.join(output_folder, "IC_Log_Summary.csv")
        # Perform IC log specific processing and save CSV to `output_folder`
        log_message(f"Unzipping IC logs completed. Summary saved at: {csv_summary_path}")
    else:
        log_message("No IC log zip file found. Skipping extraction.")

    log_message("IC log extraction process completed.")


def find_zip_file(folder_path):
    log_message(f"Searching for zip files in directory: {folder_path}")
    abs_folder_path = os.path.abspath(folder_path)
    for root, dirs, files in os.walk(abs_folder_path):
        for file in files:
            if file.endswith(".zip"):
                zip_file_path = os.path.join(root, file)
                log_message(f"Found zip file: {zip_file_path}")
                return zip_file_path
    log_message(f"No zip file found in directory: {folder_path}")
    return None


def unzip_file(zip_path, password):
    log_message(f"Unzipping file: {zip_path} with password: {password}")
    abs_zip_path = os.path.abspath(zip_path)
    unzip_dir = os.path.dirname(abs_zip_path)
    log_message(f"Output folder path: {unzip_dir}")
    try:
        with zipfile.ZipFile(abs_zip_path, "r") as zip_ref:
            total_files = len(zip_ref.infolist())
            for i, file_info in enumerate(zip_ref.infolist(), start=1):
                log_message(f"Extracting...: {file_info.filename} ({i}/{total_files})")
                zip_ref.extract(file_info, path=unzip_dir, pwd=password.encode() if password else None)
                
    except RuntimeError as e:
        log_message(f"Failed to unzip {zip_path}: {str(e)}")


def reformat_backendlogs_linefeeds(folder_path, pattern):
    # Get the list of all files matching the pattern in the specified folder
    files = [f for f in os.listdir(folder_path) if fnmatch.fnmatch(f, pattern)]
    total_files = len(files)

    # Loop through all matching files in the folder
    for index, filename in enumerate(files, start=1):
        file_path = os.path.join(folder_path, filename)
        
        # Display progress at the beginning
        log_message(f"Reformatting linefeeds for backend log file: {index}/{total_files}: {filename}")

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
        log_message("No files found matching the pattern.")
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
        log_message(f"File {i + 1} of {total_files}: {file_name}\t Start: {file_oldest_date} End: {file_youngest_date} Difference: {formatted_file_diff}")

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
        log_message(f"Overall oldest date: {overall_oldest_date}")
        log_message(f"Overall youngest date: {overall_youngest_date}")
        log_message(f"Overall difference: {formatted_diff}")
    else:
        log_message("No valid dates found in the specified files.")


def split_log_files(folder_path, partfoldername, regex_pattern, basepart):
    """
    Splits large log files into smaller parts with a max size of 20,000 KB.
    Creates a 'parts' subfolder to save the split files.
    """
    log_message(f"Splitting log files in folder: {os.path.abspath(folder_path)} into '{partfoldername}'...")    
    # Create the folder where split log files will be stored
    parts_folder = os.path.join(folder_path, partfoldername)
    if os.path.exists(parts_folder):
        shutil.rmtree(parts_folder)  # Remove existing folder if it exists
    os.makedirs(parts_folder)
    log_message(f"Created parts folder: {parts_folder}")

    # List all files that match the regex pattern
    log_files = [file for file in os.listdir(folder_path) if regex_pattern.match(file)]
    log_message(f"Pattern detected {len(log_files)} matching file(s) in folder: {folder_path}.")
    if not log_files:
        log_message(f"No files matching pattern in folder: {folder_path}")
        return

    # Process each log file and split into parts
    for log_file in log_files:
        log_file_path = os.path.join(folder_path, log_file)
        log_message(f"Processing log file: {log_file_path}")

        try:
            # Read the file content
            with open(log_file_path, "r", encoding="utf-8") as file:
                content = file.read()
            file_size = os.path.getsize(log_file_path)

            # Calculate the number of parts needed for splitting
            max_part_size = 20_000 * 1024  # 20,000 KB in bytes
            num_parts = math.ceil(file_size / max_part_size)

            if num_parts == 0:
                log_message(f"File '{log_file}' is smaller than the max size and does not need splitting. Skipping...")
                continue

            # Extract the suffix if it exists (e.g., .1, .2, etc.)
            suffix = ""
            if "." in log_file:
                suffix = "." + log_file.split(".")[-1] if log_file.split(".")[-1].isdigit() else ""  # Keep the numeric suffix
                base_name = ".".join(log_file.split(".")[:-1])  # Handle cases with multiple dots
            else:
                base_name = log_file  # No suffix

            remaining_content = content
            for i in range(num_parts):
                part = remaining_content[:max_part_size]
                remaining_content = remaining_content[max_part_size:]

                # Construct the part filename
                part_filename = f"{base_name}{suffix}.part{i + 1:03}"
                part_filepath = os.path.join(parts_folder, part_filename)

                # Save the part to a new file
                with open(part_filepath, "w", encoding="utf-8") as part_file:
                    part_file.write(part)

                log_message(f"Created part file: {part_filepath}")
        
        except Exception as e:
            log_message(f"Error processing file {log_file_path}: {str(e)}")
            continue

    log_message(f"Finished splitting logs in folder: {folder_path}")        

def save_xml_with_utf8_encoding(file_path):
    log_message(f"Saving XML file with UTF-8 encoding: {file_path}")
    try:
        with open(file_path, "r", encoding="utf-8") as file:
            content = file.read()
        with open(file_path, "w", encoding="utf-8") as file:
            file.write(content)
        log_message(f"File saved successfully with UTF-8 encoding: {file_path}")
    except Exception as e:
        log_message(f"Error saving XML file {file_path} with UTF-8 encoding: {str(e)}")


def create_list_of_files(startpath):
    """
    Creates a CSV file listing all files and folders in the directory.
    Output is written to 'Process output/problemreport_files.csv'.
    """        
    log_message(f"Creating a list of files in directory: {startpath}")
    
    # Define the output path for the file list
    output_folder = setup_output_folder(startpath)
    file_path = os.path.join(output_folder, PROBLEM_REPORT_FILE)

    with open(file_path, "w", newline="", encoding="utf-8") as csvfile:
        filewriter = csv.writer(csvfile, delimiter=";", quotechar="|", quoting=csv.QUOTE_MINIMAL)
        filewriter.writerow(["Name", "Type", "Path"])
        for root, dirs, files in os.walk(startpath):
            for name in files:
                filewriter.writerow([name, "File", os.path.join(root, name)])
            for dir_name in dirs:
                filewriter.writerow([dir_name, "Folder", os.path.join(root, dir_name)])
    log_message(f"List of files saved into: {file_path}")        

def extract_pipetting_steps(root_folder):
    """
    Executes the PipettorDataExtractorConsole.exe located in PipettorDataExtractor_Tool subfolder 
    with configured parameters. Outputs a CSV file to the 'Process output' folder.
    """    
    log_message("Starting pipetting steps extraction process.")

    # Path to the PipettorDataExtractorConsole.exe
    exe_path = os.path.join(root_folder, "PipettorDataExtractor_Tool", "PipettorDataExtractorConsole.exe")

    # Output file saved to 'Process output'
    output_folder = setup_output_folder(root_folder)  # Create/confirm the output folder
    csv_file_path = os.path.join(output_folder, PIPETTING_EVENTS_FILE)  # Define output CSV file path
    input_path = os.path.abspath(root_folder)  # Get absolute path for input folder

    # Construct the command as a LIST of arguments for subprocess.run
    command_list = [
        exe_path,
        "-p",
        input_path,        # Input folder path
        csv_file_path,     # Output CSV file in 'Process output'
        "true",            # Overwrite existing file flag
        "true",            # Extract all data flag
        "false"            # Additional flag
    ]

    log_message(f"Command list to execute: {command_list}")

    try:
        # Execute command and capture output
        result = subprocess.run(command_list, text=True, capture_output=True, check=False)
        
        # Process and log each line from stdout
        for line in result.stdout.splitlines():
            log_message(f"stdout: {line}")

        # Process and log each line from stderr
        for line in result.stderr.splitlines():
            log_message(f"stderr: {line}")
        
        # Check the return code
        if result.returncode == 0:
            log_message(f"Execution successful! CSV generated at: {csv_file_path}")
        else:
            log_message(f"Execution failed with error code {result.returncode}.")
            log_message(f"Command stderr:\n{result.stderr}")

    except FileNotFoundError:
        log_message(f"Error: Could not find the executable at: {exe_path}. Ensure it exists in the PipettorDataExtractor_Tool subfolder.")
    except Exception as e:
        log_message(f"An unexpected error occurred while executing command {command_list}: {str(e)}")    


def analyze_ebarcodes_importer(root_folder):
    """
    Analyzes e-Barcodes Importer logs from backend log parts and generates a CSV summary in 'Process output/'.
    """    
    log_message("Starting e-Barcodes importer analysis...")

    # Define the path where the log files are located
    relative_path = r"EventTraces\SystemSoftwareLog"
    log_dir = os.path.join(root_folder, relative_path)

    # Define the output folder and summary file path
    output_folder = setup_output_folder(root_folder)
    csv_summary_file = os.path.join(output_folder, EBARCODE_SUMMARY_FILE)

    # Ensure the log directory exists
    if not os.path.exists(log_dir):
        log_message(f"Error: Directory not found: {log_dir}")
        return

    log_message(f"Processing e-Barcodes logs from directory: {log_dir}")
    log_count = 0
    processed_rows = []

    def sanitize_version(version):
        """Ensure the version field contains only digits and periods (e.g., '8.2.101')."""
        return re.sub(r"[^\d.]", "", version) if version else None

    def extract_message_details(message):
        """Extract MasterFileCategory, Version, and Code from the message."""
        try:
            # Split the message and extract details
            parts = message.split(", ")
            category = next(
                (p.split(": ")[1].strip() for p in parts if "MasterFileCategory:" in p), None
            )
            version = next(
                (p.split(": ")[1].strip() for p in parts if "Version:" in p), None
            )
            code = next((p.split(": ")[1].strip() for p in parts if "Code:" in p), None)
            version = sanitize_version(version)  # Clean up the version string
            if code and "..." in code:
                code = code.replace("...", "").strip()  # Clean up ellipsis
            return category, version, code
        except Exception:
            return None, None, None

    def extract_install_events(files):
        """Extract all lines with InstallProgressChangedEvent from all given files."""
        install_events = []
        for file_name in files:
            file_path = os.path.join(log_dir, file_name)
            log_message(f"Processing install events: {file_name}")
            try:
                with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                    for line in f:
                        if "InstallProgressChangedEvent" in line:
                            install_events.append({"event": line.strip(), "file_name": file_name})
            except Exception as e:
                log_message(f"Error extracting install events from '{file_path}': {e}")
        return install_events

    def find_install_event(version, code, auxiliary_events):
        """Search the auxiliary data for an install event with matching version and code."""
        for event in auxiliary_events:
            if f'Code: "{code}"' in event["event"] and f'Version: "{version}"' in event["event"]:
                return event["event"]  # Return the matched installation event
        return None

    # Discover log files
    try:
        log_files = sorted(
            (
                f for f in os.listdir(log_dir) if f.startswith("Roche.C4C.ServiceHosting.BackEnd.log")
            ),
            key=lambda f: os.path.getctime(os.path.join(log_dir, f)),
            reverse=True,
        )
    except Exception as e:
        log_message(f"Error while fetching log files: {e}")
        return

    # Step 1: Extract auxiliary events to match with main log events
    auxiliary_events = extract_install_events(log_files)

    # Step 2: Analyze log files and generate rows for the CSV
    for file_name in log_files:
        file_path = os.path.join(log_dir, file_name)
        log_message(f"Processing: {file_name}")

        try:
            with open(file_path, "r", encoding="utf-8", errors="ignore") as log_file:
                for line_number, line in enumerate(log_file, start=1):
                    # Use delimiter `<!>` for splitting lines
                    parts = line.split("<!>")
                    if len(parts) < 7:  # Skip incomplete lines
                        continue

                    timestamp = parts[0].strip()
                    logger = parts[3].strip()
                    message = parts[4].strip()

                    # Skip irrelevant loggers
                    if logger == "":
                        continue

                    if "Roche.C4C.RSA.BusinessLogic.PackageManagement.PackageImporting" in logger:
                        master_file_category, version, code = extract_message_details(message)
                        install_event = None
                        if version and code:
                            install_event = find_install_event(version, code, auxiliary_events)

                        # Append the row information for CSV output
                        processed_rows.append([
                            timestamp,
                            master_file_category,
                            version,
                            code,
                            message,
                            install_event or "",
                        ])
                        log_count += 1
        except Exception as e:
            log_message(f"Error processing log file '{file_name}': {e}")

    # Write the final CSV file
    try:
        with open(csv_summary_file, "w", encoding="utf-8", newline="") as csv_file:
            writer = csv.writer(csv_file, delimiter=";")
            writer.writerow(["Timestamp", "MasterFileCategory", "Version", "Code", "Message", "Installation Event"])
            writer.writerows(processed_rows)
        log_message(f"e-Barcodes Importer Summary saved to: {csv_summary_file}")
    except Exception as e:
        log_message(f"Failed to write summary to '{csv_summary_file}': {e}")

    if log_count == 0:
        log_message("No relevant logs processed.")
    else:
        log_message(f"Finished processing files. Total relevant logs: {log_count}")

    log_message("e-Barcodes importer analysis completed.")    

def extract_reboot_information(root_folder):
    """
    Extracts reboot information from log files matching the pattern 'Roche.C4C.ServiceHosting.BackEnd.log.*'
    in the folder 'EventTraces/SystemSoftwareLog'. Saves the data to 'Process output/reboots.csv',
    including the date found at the beginning of each trace and the source file name as the last column.
    """    
    log_message("Starting reboot information extraction...")

    # Define the relative path and file pattern
    relative_path = r"EventTraces\SystemSoftwareLog"
    log_dir = os.path.join(root_folder, relative_path)
    file_pattern = r"Roche.C4C.ServiceHosting.BackEnd.log.*"

    # Output setup
    output_folder = setup_output_folder(root_folder)
    csv_file_path = os.path.join(output_folder, "reboots.csv")

    # Ensure the log directory exists
    if not os.path.exists(log_dir):
        log_message(f"Error: Directory not found: {log_dir}")
        return

    # Define headers with "Source File Name" as the last column and "Date" as the first column
    headers = [
        "Date", "Product Name", "Copyright Information", "Executable",
        "Assembly Version", "File Version", "Product Version", "Hardware Version",
        "Instrument Type", "Using Clean Database", "Source File Name"
    ]

    # List of extracted rows for the CSV file
    extracted_data = []

    def parse_trace_block(log_block, source_file):
        """
        Parse a single block of traces and extract the required fields into a list.
        Adds the date and source file name to the row.
        """
        data = {
            "Date": "",
            "Product Name": "",
            "Copyright Information": "",
            "Executable": "",
            "Assembly Version": "",
            "File Version": "",
            "Product Version": "",
            "Hardware Version": "",
            "Instrument Type": "",
            "Using Clean Database": "",
            "Source File Name": source_file,
        }

        # Extract fields from the block
        try:
            for line in log_block.splitlines():                           
                if "Product name:" in line:
                    data["Product Name"] = line.split(": ", 1)[1].split("<!>")[0].strip()
                elif "Copyright information:" in line:
                    data["Copyright Information"] = line.split(": ", 1)[1].split("<!>")[0].strip()
                elif "Executable:" in line:
                    data["Executable"] = line.split(": ", 1)[1].split("<!>")[0].strip()
                elif "Assembly version:" in line:
                    data["Assembly Version"] = line.split(": ", 1)[1].split("<!>")[0].strip()
                elif "File version:" in line:
                    data["File Version"] = line.split(": ", 1)[1].split("<!>")[0].strip()
                elif "Product version:" in line:
                    data["Product Version"] = line.split(": ", 1)[1].split("<!>")[0].strip()
                elif "Hardware version:" in line:
                    data["Hardware Version"] = line.split(": ", 1)[1].split("<!>")[0].strip()
                elif "Instrument Type:" in line:
                    data["Instrument Type"] = line.split(": ", 1)[1].split("<!>")[0].strip()
                elif "Using clean database:" in line:
                    data["Using Clean Database"] = line.split(": ", 1)[1].split("<!>")[0].strip()
                else:
                    potential_date = line.split("<!>")[0].strip()
                    if potential_date:  # If a valid date exists
                        data["Date"] = potential_date
            # Convert the extracted dictionary to a row format and ensure it's not empty
            row = [data[header] for header in headers]
            # Validate if any column (except 'Source File Name') has data
            if any(row[1:-1]):  # Skip 'Source File Name' (last column) and 'Date' (first column) in the emptiness check
                return row
            else:
                return None  # Skip empty rows
        except Exception as e:
            log_message(f"Error parsing trace block in file '{source_file}': {e}")
            return None

    # Process log files
    log_files = [file for file in os.listdir(log_dir) if fnmatch.fnmatch(file, file_pattern)]
    log_message(f"Found {len(log_files)} log files matching the pattern '{file_pattern}'.")

    if not log_files:
        log_message(f"No log files matching pattern '{file_pattern}' found in: {log_dir}.")
        return

    for file_name in log_files:
        file_path = os.path.join(log_dir, file_name)
        log_message(f"Processing log file: {file_name}")

        try:
            # Read and process the entire file and split at the defined blocks
            with open(file_path, "r", encoding="utf-8", errors="ignore") as log_file:
                file_content = log_file.read()
                blocks = file_content.split("########################################################################################################################")
                for block in blocks:
                    parsed_row = parse_trace_block(block, source_file=file_name)
                    if parsed_row:
                        extracted_data.append(parsed_row)
        except Exception as e:
            log_message(f"Error reading from log file '{file_name}': {e}")

    # Write the extracted data to the CSV
    try:
        with open(csv_file_path, mode="w", newline="", encoding="utf-8") as csv_file:
            writer = csv.writer(csv_file, delimiter=";")
            # Write header and rows
            writer.writerow(headers)
            writer.writerows(extracted_data)
        log_message(f"Reboot information saved to: {csv_file_path}")
    except Exception as e:
        log_message(f"Error writing to CSV file '{csv_file_path}': {e}")

    # Log completion
    if extracted_data:
        log_message(f"Reboot analysis completed. Total reboot traces found: {len(extracted_data)}")
    else:
        log_message("No reboot information found in the logs.")    
     
def find_utc_time(root_folder):
    """
    Finds the UTC time zone information from the first <Timestamp> tag in an XML file
    located in the relative folder 'EventTraces/Logging'.
    Prints out the detected UTC offset only once.
    """    
    log_message("Starting process to find UTC time from timestamps...")

    # Define the relative folder path and XML file pattern
    relative_path = r"EventTraces\Logging"
    folder_path = os.path.join(root_folder, relative_path)
    xml_file_pattern = "*.xml"  # Look for XML files in the folder

    # Ensure the folder exists
    if not os.path.exists(folder_path):
        log_message(f"Error: Folder does not exist: {folder_path}")
        return

    # Find all XML files in the folder
    xml_files = [file for file in os.listdir(folder_path) if fnmatch.fnmatch(file, xml_file_pattern)]
    
    if not xml_files:
        log_message(f"No XML files found in the folder: {folder_path}")
        return

    # Iterate through XML files to find the first <Timestamp> tag
    for xml_file in xml_files:
        xml_file_path = os.path.join(folder_path, xml_file)
        log_message(f"Processing XML file: {xml_file}")

        try:
            # Read the content of the XML file
            with open(xml_file_path, "r", encoding="utf-8") as file:
                content = file.read()
            
            # Search for the first <Timestamp> tag and extract its value
            match = re.search(r"<Timestamp>(.*?)<\/Timestamp>", content)
            
            if match:
                timestamp = match.group(1)  # Extract the full timestamp value
                # Extract the UTC offset (e.g., "-07:00")
                utc_match = re.search(r"([+-]\d{2}):\d{2}$", timestamp)
                if utc_match:
                    utc_offset = utc_match.group(1)  # Extract the UTC value
                    log_message(f"Detected UTC: {utc_offset}")
                    print(f"Detected UTC: {utc_offset}")
                    return  # Exit after printing the first UTC value
            else:
                log_message(f"No <Timestamp> tags found in file: {xml_file}")
        except Exception as e:
            log_message(f"Error processing file {xml_file}: {e}")

    log_message("No UTC offset found in any of the XML files.")    
     
def main():
    root_folder = os.getcwd()  # Automatically detect the current working directory.
    while True:
        # Display the structured menu
        print("\nMain Menu:")
        print(" ---- Extract tools ----")
        print("[1] Extract all in one: [2]/[3]/[4]/[5]")
        print("    [2] Extract problem report")        
        print("    [3] Extract audits")
        print("    [4] Extract settings")			
        print("    [5] List files")
        print("[6] Extract IC log")
        
        print("---- SW analysis tools ----")
        print("[7] Analyze e-Barcodes importer")
        print("[8] Analyze software reboots")
        print("[9] Find UTC Time")  # Moved here and updated to [9]
        
        print("---- HW analysis tools ----")
        print("[10] Extract pipetting steps")
        
        print("---- Other tools ----")
        print("[11] Split files in smaller parts")  # Updated to [11]
        
        print("[Q] Quit")

        # Wait for user input
        option = input("\nSelect the option: ").strip().lower()

        if option == "1":
            log_message("==================================================")
            print("\nRunning all options [2]/[3]/[4]/[5]...")
            extract_problem_report(root_folder)  # Option [2]
            extract_audits(root_folder)          # Option [3]
            extract_settings(root_folder)       # Option [4]
            create_list_of_files(root_folder)   # Option [5]
            print("\nAll selected options [2]/[3]/[4]/[5] completed.")
            log_message("==================================================")
        elif option == "2":
            log_message("==================================================")
            extract_problem_report(root_folder)  # Option [2]
            log_message("==================================================")
        elif option == "3":
            log_message("==================================================")
            extract_audits(root_folder)          # Option [3]
            log_message("==================================================")
        elif option == "4":
            log_message("==================================================")
            extract_settings(root_folder)        # Option [4]
            log_message("==================================================")
        elif option == "5":
            log_message("==================================================")
            create_list_of_files(root_folder)    # Option [5]
            log_message("==================================================")
        elif option == "6":
            log_message("==================================================")
            print("\nExtracting IC logs...")
            extract_ic_log(root_folder)
            print("\nIC log extraction process completed.")
            log_message("==================================================")
        elif option == "7":
            log_message("==================================================")
            print("\nAnalyzing e-Barcodes importer...")
            analyze_ebarcodes_importer(root_folder)  # Option [7]
            print("\ne-Barcodes importer analysis completed.")
            log_message("==================================================")
        elif option == "8":
            log_message("==================================================")
            print("\nAnalyzing software reboots...")
            extract_reboot_information(root_folder)  # Option [8]
            print("\nSoftware reboot analysis completed.")
            log_message("==================================================")
        elif option == "9":
            log_message("==================================================")
            print("\nFinding UTC Time...")
            find_utc_time(root_folder)  # Option [9]
            print("\nUTC time zone extraction completed.")
            log_message("==================================================")
        elif option == "10":
            log_message("==================================================")
            print("\nExtracting pipetting steps...")
            extract_pipetting_steps(root_folder)  # Option [10]
            print("\nPipetting steps extraction completed.")
            log_message("==================================================")
        elif option == "11":
            log_message("==================================================")
            print("\nSplitting files into smaller parts...")
            split_log_files_in_parts(root_folder)  # Option [11]
            print("\nFile splitting completed.")
            log_message("==================================================")
        elif option == "q":
            print("Exiting the program...")
            break
        else:
            print("Invalid option. Please try again.")



if __name__ == "__main__":
    main()
