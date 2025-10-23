import os
import sys
import subprocess

def create_executable(script_path):
    # Ensure the script exists
    if not os.path.exists(script_path):
        print(f"Error: The script '{script_path}' does not exist.")
        return

    # Ensure PyInstaller is installed
    try:
        import PyInstaller
    except ImportError:
        print("PyInstaller is not installed. Installing it now...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])

    # Generate the executable
    print("Creating the executable...")
    try:
        subprocess.run(['pyinstaller', '--onefile', script_path], check=True)
        print("\nExecutable created successfully!")
        print("You can find the executable in the 'dist' folder.")
    except Exception as e:
        print(f"An error occurred while creating the executable: {e}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python create_executable.py <path_to_python_script>")
    else:
        script_path = sys.argv[1]
        create_executable(script_path)
