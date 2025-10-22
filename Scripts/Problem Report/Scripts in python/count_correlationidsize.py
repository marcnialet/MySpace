import os
import glob
import msvcrt  # Import for waiting for a key press

def process_files(folder, pattern):
    # Construct the search pattern
    search_pattern = os.path.join(folder, pattern)
    
    # Find all files matching the pattern
    files = glob.glob(search_pattern)
    
    # Process each file
    for file in files:
        char_count = 0
        total_line_chars = 0
        with open(file, 'r', encoding='utf-8') as f:
            for line in f:
                total_line_chars += len(line)
                if '<!>' in line and '[EOL]' in line:
                    end_index = line.index('[EOL]')
                    start_index = line.rindex('<!>', 0, end_index) + 3
                    char_count += end_index - start_index
        
        percentage = (char_count / total_line_chars) * 100 if total_line_chars > 0 else 0
        print(f"Correlation ID for file {os.path.basename(file)} has {char_count} chars ({percentage:.2f}%) from total chars {total_line_chars}")
        
    
    # Wait for user input to continue
    print("\nPress any key to continue...")
    msvcrt.getch()

# Example usage
folder = r"EventTraces\SystemSoftwareLog"
pattern = "Roche.C4C.ServiceHosting.BackEnd.log*.*"
process_files(folder, pattern)