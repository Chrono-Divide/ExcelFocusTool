# ExcelFocusTool

🚀 **ExcelFocusTool** is a lightweight and efficient **Excel focus management tool** built with **WinForms + C#**. It scans a specified folder for Excel files, verifies that the active worksheet is the first one and the focus is set to cell A1, and can automatically fix any focus issues found.

## 🔥 Features

✅ **Recursive Folder Scan** - Search through folders (and subfolders) to locate Excel files  
✅ **Focus Validation** - Check that the active worksheet is the first sheet and the focus (active cell, scroll position) is on cell A1  
✅ **Automatic Correction** - Adjust focus settings automatically when discrepancies are detected  
✅ **User-friendly Interface** - Simple and intuitive Windows Forms UI for selecting folders and displaying logs  
✅ **Detailed Logging** - Displays processing details for each file, including errors and fix status

---

## 📌 How to Use

### 1️⃣ **Download & Run**

- **Download the latest version** from the **[Releases](https://github.com/Chrono-Divide/ExcelFocusTool/releases)** page.  
- Extract the ZIP file and run `ExcelFocusTool.exe`.

### 2️⃣ **Usage Instructions**

1. **Select Folder:**  
   Click **"Select Folder"** or drag-and-drop a folder/file onto the interface to choose the directory containing your Excel files.
2. **Scan Files:**  
   The tool will automatically scan the selected folder and list each Excel file along with its focus status.
3. **Process Files:**  
   If any files have focus issues, click **"Process Files"** to automatically set the focus correctly (i.e., activate the first worksheet and set the focus to cell A1).
4. **Review Log:**  
   Check the on-screen log for detailed feedback on each file's processing status.

---

## 🛠 Requirements

- **Operating System:** Windows  
- **Microsoft Office:** Excel must be installed  
- **.NET Framework:** 4.0 or later  
- **Platform Target:** Ensure the build configuration (x86, x64, or AnyCPU) matches your installed Office version

---

## 🔧 Troubleshooting

- **COM Casting Errors:**  
  Make sure your project's platform target (x86/x64) is consistent with your installed Office. If errors persist, try re-registering Excel by running:
  ```bash
  excel /regserver
