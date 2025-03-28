# ExcelFocusTool

ExcelFocusTool is a .NET Windows Forms application designed to check and fix Excel file focus issues. The tool scans Excel files within a selected folder, verifies if the active worksheet and cell focus are correctly set (i.e., the active sheet is the first one and the focus is on cell A1), and automatically fixes the issue when necessary.

## Features

- **Scan Excel Files:** Recursively scan a folder for Excel files.
- **Validate Focus:** Check if the active worksheet and focus (active cell, scroll position) are properly set to the first worksheet and cell A1.
- **Automatic Fix:** Adjust focus settings for worksheets that do not comply.
- **Logging:** Display detailed logs for each file processed, indicating whether it is valid or what issues were found.

## Requirements

- **Operating System:** Windows
- **Microsoft Office:** Excel must be installed.
- **.NET Framework:** 4.0 or later.
- **Platform Target:** Ensure your project is built with the correct target (x86 or x64) matching your Office installation.

## Setup and Usage

1. **Clone the Repository:**
   
   ```bash
   git clone https://github.com/Chrono-Divide/ExcelFocusTool.git
   ```
