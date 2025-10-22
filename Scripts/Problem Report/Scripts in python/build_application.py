
import subprocess

def run_pyinstaller():
    command = "pyinstaller --onefile process_problem_report.py"
    try:
        subprocess.run(command, shell=True, check=True)
        print("PyInstaller command executed successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Error: {e}")
        
run_pyinstaller()
