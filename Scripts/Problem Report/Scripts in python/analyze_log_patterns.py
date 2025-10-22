
import os
import re

FOLDER_PATH = "/path/to/your/folder"  # Replace with the actual path to your folder

def find_repetitive_patterns(line):
    """
    Find repetitive patterns in a line of text.
    """
    pattern = re.compile(r'(\b\w+\b)(?=.*\b\1\b)')
    matches = pattern.findall(line)
    return matches

def analyze_files_in_folder(folder_path):
    """
    Analyze all text files in a folder for repetitive patterns.
    """
    for filename in os.listdir(folder_path):
        if filename.endswith(".txt"):
            file_path = os.path.join(folder_path, filename)
            with open(file_path, 'r') as file:
                print(f"Analyzing {filename}...")
                for line_number, line in enumerate(file, 1):
                    patterns = find_repetitive_patterns(line)
                    if patterns:
                        print(f"Line {line_number}: {line.strip()}")
                        print(f"Repetitive patterns: {patterns}")

if __name__ == "__main__":
    analyze_files_in_folder(FOLDER_PATH)
