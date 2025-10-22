
@echo off
setlocal

REM Set the hardcoded file path
set file_path=".\ProblemReport-20240508_144243158-0819f475-ac62-43c8-99bc-b88f93ef328e.zip"

REM Set the path to the Python interpreter
set python_path=python

REM Set the path to the unzipfile.py script
set script_path=process_problem_report.py

REM Execute the Python script with the hardcoded file path
%python_path% %script_path% "%file_path%"

endlocal
