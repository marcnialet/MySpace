
import os
import sys
import zipfile

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

# Check if the zip file path is provided as a command-line argument
if len(sys.argv) < 2:
    print("Please provide the path to the zip file.")
    sys.exit(1)

# Get the zip file path from the command-line argument
zip_file = sys.argv[1]

# Set the password if needed
password = "c0m1t@Pr0bl3mR3p0rt"  # Set password to an empty string for no password

# Call the unzip_file function
unzip_file(zip_file, password)
