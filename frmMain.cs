using System;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices; // ★ 新增：釋放 COM 用
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFocusTool
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterScreen; // Center on open
            this.FormBorderStyle = FormBorderStyle.FixedDialog;  // Fixed size
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

            MessageBox.Show("Please drag and drop a valid folder or file.",
                "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void btnProcessFiles_Click(object sender, EventArgs e)
        {
            string folderPath = txtFolderPath.Text;
            if (Directory.Exists(folderPath))
            {
                txtLog.AppendText("Processing started...\n");
                ProcessExcelFiles(folderPath);
                txtLog.AppendText("Processing complete.\n");
                MessageBox.Show("Processing complete!", "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Please enter a valid folder path.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 搜尋並顯示該資料夾下所有 Excel 檔案的「焦點」狀態。
        /// </summary>
        private void LogExcelFiles(string folderPath)
        {
            txtLog.Clear();

            Excel.Application excelApp = null;
            try
            {
                // 初始化 Excel COM 物件
                excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;

                // 取得所有檔案(包含子資料夾)
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

                txtLog.AppendText($"\nSummary: {validCount} files valid, {invalidCount} files invalid.\n");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel 初始化或檔案掃描時發生錯誤：\n" + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 結束 Excel 並釋放
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
            }
        }

        /// <summary>
        /// 對資料夾下所有 Excel 檔案進行「修正」(將焦點調整到 A1)。
        /// </summary>
        private void ProcessExcelFiles(string folderPath)
        {
            Excel.Application excelApp = null;
            try
            {
                excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;

                var excelFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories);
                foreach (string filePath in excelFiles)
                {
                    if (IsExcelFile(filePath))
                    {
                        string status = CheckExcelFileFocus(excelApp, filePath);
                        // 只有不在 A1 的檔案才進行修正
                        if (status != "Valid")
                        {
                            AppendLog(filePath, "Fixing...", "Blue");
                            FixExcelFile(excelApp, filePath);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel 初始化或修正時發生錯誤：\n" + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
            }
        }

        /// <summary>
        /// 檢查單一 Excel 檔案的「焦點」是否在第一張工作表、A1 位置。
        /// </summary>
        private string CheckExcelFileFocus(Excel.Application excelApp, string filePath)
        {
            Excel.Workbook workbook = null;
            try
            {
                // 只讀方式開啟
                workbook = excelApp.Workbooks.Open(filePath, ReadOnly: true);

                // 檢查當前活頁簿的「活動工作表」是否是第一張
                Excel.Worksheet activeSheet = (Excel.Worksheet)workbook.ActiveSheet;
                Excel.Worksheet firstSheet = (Excel.Worksheet)workbook.Sheets[1];

                if (activeSheet.Index != firstSheet.Index)
                {
                    return $"Active Sheet is '{activeSheet.Name}', not the first Sheet '{firstSheet.Name}'";
                }

                // 檢查所有工作表的焦點 & 滾動位置
                foreach (Excel.Worksheet worksheet in workbook.Sheets)
                {
                    int scrollRow = worksheet.Application.ActiveWindow.ScrollRow;
                    int scrollColumn = worksheet.Application.ActiveWindow.ScrollColumn;
                    Excel.Range activeCell = worksheet.Application.ActiveCell;

                    // 確認是否都是在 A1、並且滾動條在最上方/最左方
                    if (!(activeCell.Row == 1 && activeCell.Column == 1
                          && scrollRow == 1 && scrollColumn == 1))
                    {
                        return $"Sheet '{worksheet.Name}' focus => {GetCellAddress(activeCell.Row, activeCell.Column)} | ScrollRow: {scrollRow}, ScrollColumn: {scrollColumn}";
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
                // 關閉並釋放
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }
            }
        }

        /// <summary>
        /// 將 Excel 焦點修正到第一張工作表的 A1。
        /// </summary>
        private void FixExcelFile(Excel.Application excelApp, string filePath)
        {
            Excel.Workbook workbook = null;
            try
            {
                workbook = excelApp.Workbooks.Open(filePath);
                Excel.Worksheet firstSheet = (Excel.Worksheet)workbook.Sheets[1];

                // 逐張工作表處理
                foreach (Excel.Worksheet worksheet in workbook.Sheets)
                {
                    worksheet.Activate();
                    // Goto(Cells[1,1]) = 移動焦點到 A1
                    worksheet.Application.Goto(worksheet.Cells[1, 1], true);
                    worksheet.Application.ActiveWindow.ScrollRow = 1;
                    worksheet.Application.ActiveWindow.ScrollColumn = 1;
                }

                // 最後再把第一張表設為 Active
                firstSheet.Activate();

                // 存檔
                workbook.Save();
                AppendLog(filePath, "Fixed successfully", "Green");
            }
            catch (Exception ex)
            {
                AppendLog(filePath, $"Failed to fix: {ex.Message}", "Red");
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }
            }
        }

        /// <summary>
        /// 在 txtLog 顯示不同顏色的訊息
        /// </summary>
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

        /// <summary>
        /// 判斷檔案副檔名是否為 Excel
        /// </summary>
        private bool IsExcelFile(string filePath)
        {
            string ext = Path.GetExtension(filePath)?.ToLower();
            return ext == ".xls" || ext == ".xlsx" || ext == ".xlsm" || ext == ".xlsb";
        }

        /// <summary>
        /// 將 (row,col) 轉成 Excel A1 形式字串
        /// </summary>
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
