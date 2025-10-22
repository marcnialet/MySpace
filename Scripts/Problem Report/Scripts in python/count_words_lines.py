
import os
import fnmatch
import msvcrt

def count_lines_and_occurrences(folder_path, pattern, search_sentence):
    total_lines = 0
    total_occurrences = 0
    total_chars = 0
    total_founded_chars = 0
    total_percent = 0

    # Get the list of all files matching the pattern in the specified folder
    files = [f for f in os.listdir(folder_path) if fnmatch.fnmatch(f, pattern)]
    total_files = len(files)

    # Loop through all matching files in the folder
    for index, filename in enumerate(files, start=1):
        file_path = os.path.join(folder_path, filename)

        # Display progress at the beginning
        # print(f"Processing file {index}/{total_files}: {filename}")

        # Initialize counters for the current file
        line_count = 0
        occurrence_count = 0
        file_chars = 0
        founded_chars = 0
        percent = 0
        # Read the content of the file line by line
        with open(file_path, 'r', encoding='utf-8') as file:
            for line in file:
                line_count += 1
                file_chars += len(line)
                if search_sentence in line:
                    occurrence_count += 1
                    founded_chars += len(line)
        percent = (100.0 * occurrence_count) / line_count if line_count > 0 else 0
        char_percent = (100.0 * founded_chars) / file_chars if file_chars > 0 else 0

        # Update the total counters
        total_lines += line_count
        total_occurrences += occurrence_count
        total_chars += file_chars
        total_founded_chars += founded_chars
        total_percent = (100.0 * total_occurrences) / total_lines if total_lines > 0 else 0

        # Display the count for the current file
        print(f"{index}/{total_files}: {filename}: {percent:.1f} %, {line_count} lines, {occurrence_count} occurrences of '{search_sentence}' in char percent: {char_percent:.1f}%")

    # Calculate the overall percentage of repetitions over the total chars of the file
    overall_char_percent = (100.0 * total_founded_chars) / total_chars if total_chars > 0 else 0

    # Display the total counts
    print(f"Total: {total_percent:.1f}% {total_lines} lines, {total_occurrences} occurrences of '{search_sentence}'Length in char percent: {overall_char_percent:.1f}%")
    

# Specify the folder path, pattern, and search sentence here
folder_path = '.\\EventTraces\\SystemSoftwareLog'
pattern = 'Roche.C4C.ServiceHosting.BackEnd*'
search_sentence = 'Sample dilution calculation case'

count_lines_and_occurrences(folder_path, pattern, search_sentence)

print("Press any key to continue")
msvcrt.getch()

# Some statistics
# Created application description                           40% in some files 11% in CN 1624 (china)
# Sample dilution calculation case                          8.5% in some files 3% in CN 1624 (china)
# LockForPipetting (Hash:                                   9% in some files 6.7% in CN 1624 (china)
# Database command is longer then 10000 characters          4% in some files 2.4% in CN 1624 (china) (long trace)
# Event with accumulator                                    5.5% in some files 3.5% in CN 1624 (china)
# was released by LockForPipetting (Hash:                   1.5% in some files 1% in CN 1624 (china)
# HandleCycleJobResponse finished                           1% in some files 1% in CN 1624 (china)
# CanHandleJobResult returned true for TestOrderJobResponseHandlingStrategyToHandleSuccessfulResults 1% in some files 1% in CN 1624 (china)

# Total: 11.0% 2391856 lines, 261945 occurrences of 'Created application description'Length in char percent: 19.8%
# Total: 3.0%  2391856 lines, 72775 occurrences of 'Sample dilution calculation case'Length in char percent: 1.4%
# Total: 5.7%  2391856 lines, 135513 occurrences of 'LockForPipetting (Hash:'Length in char percent: 3.7%
# Total: 2.4%  2391856 lines, 56957 occurrences of 'Database command is longer then 10000 characters'Length in char percent: 1.6%
