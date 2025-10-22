
import os
import glob
from datetime import datetime

try:
    import msvcrt
    def press_any_key():
        print("Press any key to continue...")
        msvcrt.getch()
except ImportError:
    def press_any_key():
        input("Press Enter to continue...")

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

# Define the path and pattern variables
log_path = ".\\EventTraces\\SystemSoftwareLog"
file_pattern = "Roche.C4C.ServiceHosting.BackEnd*.*"

# Call the function with the defined variables
find_dates_in_files(log_path, file_pattern)

# Wait for user input to continue
press_any_key()
