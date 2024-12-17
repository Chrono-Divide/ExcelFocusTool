using System;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFocusTool
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterScreen; // Center on open
            this.FormBorderStyle = FormBorderStyle.FixedDialog; // Fixed size
            this.MaximizeBox = false; // Disable maximize
            this.MinimizeBox = true;  // Allow minimize
        }

        private void btnSelectFolder_Click(object sender, EventArgs e)
        {
            using (var folderBrowserDialog = new FolderBrowserDialog())
            {
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    txtFolderPath.Text = folderBrowserDialog.SelectedPath;
                    LogExcelFiles(folderBrowserDialog.SelectedPath);
                }
            }
        }

        private void txtFolderPath_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy; // Allow dragging
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void txtFolderPath_DragDrop(object sender, DragEventArgs e)
        {
            string[] paths = (string[])e.Data.GetData(DataFormats.FileDrop);

            foreach (string path in paths)
            {
                if (Directory.Exists(path))
                {
                    txtFolderPath.Text = path;
                    txtLog.AppendText($"Folder Path Selected: {path}\n");
                    LogExcelFiles(path);
                    return;
                }
                else if (File.Exists(path))
                {
                    string folderPath = Path.GetDirectoryName(path);
                    txtFolderPath.Text = folderPath;
                    txtLog.AppendText($"File Path Selected, Folder: {folderPath}\n");
                    LogExcelFiles(folderPath);
                    return;
                }
            }

            MessageBox.Show("Please drag and drop a valid folder or file.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void btnProcessFiles_Click(object sender, EventArgs e)
        {
            string folderPath = txtFolderPath.Text;
            if (Directory.Exists(folderPath))
            {
                txtLog.AppendText("Processing started...\n");
                ProcessExcelFiles(folderPath);
                txtLog.AppendText("Processing complete.\n");
                MessageBox.Show("Processing complete!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Please enter a valid folder path.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LogExcelFiles(string folderPath)
        {
            txtLog.Clear();
            var excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;

            var excelFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories);
            int validCount = 0, invalidCount = 0;

            foreach (string filePath in excelFiles)
            {
                if (IsExcelFile(filePath))
                {
                    string status = CheckExcelFileFocus(excelApp, filePath);
                    if (status == "Valid")
                    {
                        AppendLog(filePath, "Focus is on A1", "Green");
                        validCount++;
                    }
                    else
                    {
                        AppendLog(filePath, status, "Red");
                        invalidCount++;
                    }
                }
            }

            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            txtLog.AppendText($"\nSummary: {validCount} files valid, {invalidCount} files invalid.\n");
        }

        private void ProcessExcelFiles(string folderPath)
        {
            var excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;

            var excelFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories);
            foreach (string filePath in excelFiles)
            {
                if (IsExcelFile(filePath))
                {
                    string status = CheckExcelFileFocus(excelApp, filePath);
                    if (status != "Valid") // Only process invalid files
                    {
                        AppendLog(filePath, "Fixing...", "Blue");
                        FixExcelFile(excelApp, filePath);
                    }
                }
            }

            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }

        private string CheckExcelFileFocus(Excel.Application excelApp, string filePath)
        {
            Excel.Workbook workbook = null;
            try
            {
                workbook = excelApp.Workbooks.Open(filePath, ReadOnly: true);

                // 检查当前活动的工作表是否是第一个工作表（最左边）
                Excel.Worksheet activeSheet = (Excel.Worksheet)workbook.ActiveSheet;
                Excel.Worksheet firstSheet = (Excel.Worksheet)workbook.Sheets[1];

                if (activeSheet.Index != firstSheet.Index)
                {
                    return $"Active Sheet is '{activeSheet.Name}', not the first Sheet '{firstSheet.Name}'";
                }

                // 检查所有工作表的焦点和滚动位置
                foreach (Excel.Worksheet worksheet in workbook.Sheets)
                {
                    int scrollRow = worksheet.Application.ActiveWindow.ScrollRow;
                    int scrollColumn = worksheet.Application.ActiveWindow.ScrollColumn;

                    Excel.Range activeCell = worksheet.Application.ActiveCell;

                    // 判断是否满足焦点和滚动位置的条件
                    if (!(activeCell.Row == 1 && activeCell.Column == 1 && scrollRow == 1 && scrollColumn == 1))
                    {
                        return $"Sheet '{worksheet.Name}' focus is on {GetCellAddress(activeCell.Row, activeCell.Column)} | ScrollRow: {scrollRow}, ScrollColumn: {scrollColumn}";
                    }
                }

                return "Valid";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
            finally
            {
                workbook?.Close(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            }
        }


        private void FixExcelFile(Excel.Application excelApp, string filePath)
        {
            Excel.Workbook workbook = null;
            try
            {
                workbook = excelApp.Workbooks.Open(filePath);
                Excel.Worksheet firstSheet = (Excel.Worksheet)workbook.Sheets[1]; // Keep track of first sheet

                foreach (Excel.Worksheet worksheet in workbook.Sheets)
                {
                    worksheet.Activate();
                    worksheet.Application.Goto(worksheet.Cells[1, 1], true);
                    worksheet.Application.ActiveWindow.ScrollRow = 1;
                    worksheet.Application.ActiveWindow.ScrollColumn = 1;
                }

                // Activate the first sheet after processing all sheets
                firstSheet.Activate();
                workbook.Save();
                AppendLog(filePath, "Fixed successfully", "Green");
            }
            catch (Exception ex)
            {
                AppendLog(filePath, $"Failed to fix: {ex.Message}", "Red");
            }
            finally
            {
                workbook?.Close(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            }
        }

        private void AppendLog(string filePath, string message, string color)
        {
            txtLog.SelectionStart = txtLog.TextLength;
            txtLog.SelectionLength = 0;

            switch (color)
            {
                case "Red":
                    txtLog.SelectionColor = System.Drawing.Color.Red;
                    break;
                case "Green":
                    txtLog.SelectionColor = System.Drawing.Color.Green;
                    break;
                case "Blue":
                    txtLog.SelectionColor = System.Drawing.Color.Blue;
                    break;
            }

            txtLog.AppendText($"{filePath}: {message}\n");
            txtLog.SelectionStart = txtLog.TextLength;
            txtLog.ScrollToCaret();
            txtLog.SelectionColor = txtLog.ForeColor;
        }

        private bool IsExcelFile(string filePath)
        {
            string ext = Path.GetExtension(filePath)?.ToLower();
            return ext == ".xls" || ext == ".xlsx" || ext == ".xlsm" || ext == ".xlsb";
        }

        private string GetCellAddress(int row, int column)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string address = "";
            while (column > 0)
            {
                int index = (column - 1) % 26;
                address = letters[index] + address;
                column = (column - 1) / 26;
            }
            return address + row;
        }
    }
}
