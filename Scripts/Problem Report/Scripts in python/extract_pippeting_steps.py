import os
import subprocess

def main():
    # Step 1: Automatically determine the root directory where the script is executed
    root_folder = os.getcwd()

    # Step 2: Construct the paths for the executable and the output CSV file
    exe_path = os.path.join(root_folder, "PipettorDataExtractorConsole.exe")
    csv_file_path = os.path.join(root_folder, "pipetting_logs.csv")

    # Step 3: Define additional parameters
    parameter_p = f'-p "{root_folder}\\"'  # Path to root folder (with trailing backslash)
    overwrite_existing = "true"  # Replace CSV file if it exists
    extract_all = "true"
    additional_param = "false"

    # Print the exe_path and parameters before building the command
    print(f"Executable Path: {exe_path}")
    print(f"Parameter -p: {parameter_p}")
    print(f"CSV File Path: {csv_file_path}")
    print(f"Additional Parameters: {overwrite_existing}, {extract_all}, {additional_param}")

    # Step 4: Build the command
    command = f'"{exe_path}" {parameter_p} {csv_file_path} {overwrite_existing} {extract_all} {additional_param}'

    try:
        # Step 5: Execute the command
        result = subprocess.run(command, shell=True, capture_output=True, text=True)

        # Step 6: Output success or error message
        if result.returncode == 0:
            print(f"Execution successful! CSV generated at: {csv_file_path}")
        else:
            print(f"Execution failed with error code {result.returncode}")
            print(f"Error Output: {result.stderr}")

    except FileNotFoundError:
        print(f"Error: Could not find executable at: {exe_path}")
    except Exception as e:
        print(f"An unexpected error occurred: {str(e)}")

if __name__ == "__main__":
    main()
