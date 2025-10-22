import os
import csv
import re

# Define the relative path where the log files are located
relative_path = r".\EventTraces\SystemSoftwareLog\parts_backendlog"
log_dir = os.path.join(os.getcwd(), relative_path)

# Output file name
csv_summary_file = "eBarcode Importer Summary.csv"

# Ensure the directory exists
if not os.path.exists(log_dir):
    print(f"Error: Directory not found: {log_dir}")
    sys.exit(1)

# Trace summary storage
log_count = 0

# Define the logger to exclude
excluded_logger = (
    #"Roche.C4C.RSA.BusinessLogic.PackageManagement.PackageImporting."
    #"Engines.EBarcodePackageImportCoordinator[2_5_0_2881]"
)

def sanitize_version(version):
    """Ensure the version field contains only digits and period (e.g., '8.2.101')."""
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
            code = code.replace("...", "").strip()  # Clean up ellipses
        return category, version, code
    except Exception:
        return None, None, None

def extract_install_events(files):
    """Extract all lines with InstallProgressChangedEvent from all given files."""
    install_events = []
    for file_name in files:
        print(f"Processing install events: {file_name}")
        file_path = os.path.join(log_dir, file_name)
        try:
            with open(file_path, "r", encoding="utf-8", errors="ignore") as file:
                for line in file:
                    if "InstallProgressChangedEvent" in line:
                        install_events.append({"event": line.strip(), "file_name": file_name})  # Add matching lines to the list
        except Exception as e:
            print(f"DEBUG: Error extracting install events from file {file_path}: {e}")
    return install_events

def find_install_event(version, code, auxiliary_events):
    """Search the auxiliary data for an install event with matching version and code."""
    for event in auxiliary_events:
        if f'Code: "{code}"' in event["event"] and f'Version: "{version}"' in event["event"]:
            return event["event"]  # Return the matched installation event
    return None

print("\nProcessing log files...\n")

# Get all log files sorted by creation date (descending)
log_files = sorted(
    (
        f
        for f in os.listdir(log_dir)
        if f.startswith("Roche.C4C.ServiceHosting.BackEnd.log")
    ),
    key=lambda f: os.path.getctime(os.path.join(log_dir, f)),
    reverse=True,
)

# Step 1: Extract installation events from files
print("\nExtracting installation events from log files...\n")
auxiliary_events = extract_install_events(log_files)

# Step 2: Process log files and generate eBarcode Importer Summary
processed_rows = []

for file_name in log_files:
    file_path = os.path.join(log_dir, file_name)
    print(f"Processing: {file_name}")

    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as file:
            for line_number, line in enumerate(file, start=1):
                # Split the line manually using the delimiter `<!>`
                parts = line.split("<!>")
                if len(parts) < 7:  # Skip incomplete lines
                    continue

                timestamp = parts[0].strip()
                logger = parts[3].strip()
                message = parts[4].strip()

                # Skip the specified logger
                if logger == excluded_logger:
                    continue

                if "Roche.C4C.RSA.BusinessLogic.PackageManagement.PackageImporting" in logger:
                    # Extract MasterFileCategory, Version, and Code from the message
                    master_file_category, version, code = extract_message_details(message)

                    # Search for an installation event in the auxiliary data
                    install_event = None
                    if version and code:
                        install_event = find_install_event(version, code, auxiliary_events)

                    # Append the row information for CSV output in the specified order
                    processed_rows.append([
                        timestamp,
                        master_file_category,
                        version,
                        code,
                        message,
                        install_event if install_event else "",
                    ])
                    log_count += 1
    except Exception as e:
        print(f"Error processing file {file_name}: {e}")

if log_count == 0:
    print("No relevant logs processed.")
else:
    print(f"\nFinished processing files. Total relevant logs: {log_count}")

# Write to CSV file
try:
    with open(csv_summary_file, "w", encoding="utf-8", newline="") as csv_file:
        writer = csv.writer(csv_file, delimiter=";")
        # Write the header in the specified order
        writer.writerow(["Timestamp", "MasterFileCategory", "Version", "Code", "Message", "Installation Event"])
        # Write all processed rows
        for row in processed_rows:
            writer.writerow(row)

    print(f"\nFinal summary saved to: {csv_summary_file}")
except Exception as e:
    print(f"Error writing to {csv_summary_file}: {e}")
