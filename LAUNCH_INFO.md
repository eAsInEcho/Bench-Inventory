# IT Bench Inventory Manager - Launch Information

## Issue: Application Not Launching After Visual Studio Installation

After installing Visual Studio 2022 with Python workload, the application stopped launching properly when double-clicking `main.py`. The command window would flash briefly and close without the application starting.

## Why This Happens

When Visual Studio with Python workload is installed, it can modify how Python files (`.py`) are executed when double-clicked:

1. The file association for `.py` files may have changed
2. The Python interpreter runs the script but immediately closes the console window after completion or if errors occur
3. This prevents any error messages from being visible and stops the GUI from initializing properly

## Solution: Batch File Launchers

Two batch files have been created to help launch the application properly:

### 1. `launch_app.bat` (Simple Version)

This batch file:
- Runs the main.py script using the default Python installation
- Keeps the console window open if any errors occur
- Displays the error code for troubleshooting

### 2. `launch_app_alternative.bat` (Advanced Version)

This more robust batch file:
- First attempts to use Python from the system PATH
- If not found, tries common Python installation locations
- Provides clear feedback about which Python installation is being used
- Keeps the console window open if any errors occur
- Displays helpful error messages

## How to Use

1. Double-click on either `launch_app.bat` or `launch_app_alternative.bat` to start the application
2. If one doesn't work, try the other
3. If errors occur, the console window will remain open so you can read the error messages

## Troubleshooting

If neither batch file works:

1. Check that Python is installed correctly
2. Ensure all required packages are installed (`pip install -r requirements.txt`)
3. Try editing the batch file to point to the correct Python installation path on your system
4. Check for any error messages in the console window
