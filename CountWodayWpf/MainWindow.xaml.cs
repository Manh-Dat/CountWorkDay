using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using ClosedXML.Excel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;

namespace CountWodayWpf
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private const string SaveFilePath = "user_rows.txt";

        public MainWindow()
        {
            // Ensure MainWindow.xaml exists and is set to Build Action: Page
            InitializeComponent();
        }

        private void SelectFileButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == true)
            {
                InputFilePathTextBox.Text = openFileDialog.FileName;
            }
        }

        private void SelectOutputFileButton_Click(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new Microsoft.Win32.SaveFileDialog();
            saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            saveFileDialog.FileName = "output.xlsx";
            if (saveFileDialog.ShowDialog() == true)
            {
                OutputFilePathTextBox.Text = saveFileDialog.FileName;
            }
        }

        public void AppendDebug(string message)
        {
            DebugTextBox.AppendText(message + "\n");
            DebugTextBox.ScrollToEnd();
        }

        private string GetExcelColumnName(int columnNumber)
        {
            string columnName = "";
            while (columnNumber >= 0)
            {
                columnName = (char)('A' + (columnNumber % 26)) + columnName;
                columnNumber = columnNumber / 26 - 1;
            }
            return columnName;
        }

        private void ProcessButton_Click(object sender, RoutedEventArgs e)
        {
            DebugTextBox.Clear();
            string inputPath = InputFilePathTextBox.Text;
            if (string.IsNullOrWhiteSpace(inputPath) || !System.IO.File.Exists(inputPath))
            {
                AppendDebug("Vui lòng chọn file Excel input hợp lệ.");
                return;
            }
            try
            {
                IWorkbook workbook;
                using (var fs = new System.IO.FileStream(inputPath, System.IO.FileMode.Open, System.IO.FileAccess.Read))
                {
                    if (inputPath.EndsWith(".xls"))
                        workbook = new HSSFWorkbook(fs);
                    else if (inputPath.EndsWith(".xlsx"))
                        workbook = new XSSFWorkbook(fs);
                    else
                        throw new Exception("Định dạng file không hỗ trợ");

                    AppendDebug($"Đã mở file: {inputPath}");
                    for (int i = 0; i < workbook.NumberOfSheets; i++)
                    {
                        var sheet = workbook.GetSheetAt(i);
                        bool hasEmployeeTable = false;
                        for (int rowIdx = 0; rowIdx <= sheet.LastRowNum && !hasEmployeeTable; rowIdx++)
                        {
                            var row = sheet.GetRow(rowIdx);
                            if (row == null) continue;
                            for (int colIdx = 0; colIdx < row.LastCellNum; colIdx++)
                            {
                                var cell = row.GetCell(colIdx);
                                if (cell != null && cell.CellType == CellType.String && cell.StringCellValue.Trim() == "Employee Attendance Table")
                                {
                                    hasEmployeeTable = true;
                                    break;
                                }
                            }
                        }
                        if (!hasEmployeeTable) continue;
                        AppendDebug($"--- Sheet: {sheet.SheetName} ---");
                        for (int rowIdx = 0; rowIdx <= sheet.LastRowNum; rowIdx++)
                        {
                            var row = sheet.GetRow(rowIdx);
                            if (row == null) continue;
                            for (int colIdx = 0; colIdx < row.LastCellNum - 1; colIdx++)
                            {
                                var cell = row.GetCell(colIdx);
                                if (cell != null && cell.CellType == CellType.String && cell.StringCellValue.Trim() == "Name")
                                {
                                    var nameCell = row.GetCell(colIdx + 1);
                                    string empName = nameCell != null ? nameCell.ToString().Trim() : "";
                                    if (!string.IsNullOrEmpty(empName))
                                    {
                                        // Tìm dòng chứa "Time Card" bên dưới ô Name
                                        int timeCardRow = -1;
                                        int timeCardColStart = colIdx;
                                        int timeCardColEnd = colIdx;
                                        for (int r = rowIdx + 1; r <= sheet.LastRowNum; r++)
                                        {
                                            var checkRow = sheet.GetRow(r);
                                            if (checkRow == null) continue;
                                            for (int c = colIdx; c < checkRow.LastCellNum; c++)
                                            {
                                                var checkCell = checkRow.GetCell(c);
                                                if (checkCell != null && checkCell.CellType == CellType.String && checkCell.StringCellValue.Trim().Contains("Time Card"))
                                                {
                                                    timeCardRow = r;
                                                    // Xác định dải cột của dòng Time Card
                                                    // Tìm cột bắt đầu
                                                    timeCardColStart = c;
                                                    // Tìm cột kết thúc (liên tiếp có dữ liệu)
                                                    timeCardColEnd = c;
                                                    for (int tc = c + 1; tc < checkRow.LastCellNum; tc++)
                                                    {
                                                        var nextCell = checkRow.GetCell(tc);
                                                        if (nextCell != null && !string.IsNullOrWhiteSpace(nextCell.ToString()))
                                                            timeCardColEnd = tc;
                                                        else
                                                            break;
                                                    }
                                                    break;
                                                }
                                            }
                                            if (timeCardRow != -1) break;
                                        }
                                        if (timeCardRow != -1)
                                        {
                                            string startCell = GetExcelColumnName(timeCardColStart) + (timeCardRow + 1);
                                            string endCell = GetExcelColumnName(timeCardColEnd) + (timeCardRow + 1);
                                            AppendDebug($"Nhân viên: {empName} | Dải ô Time Card: {startCell} - {endCell}");
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                AppendDebug($"Lỗi khi đọc file: {ex.Message}");
            }
        }

        private void AddRowButton_Click(object sender, RoutedEventArgs e)
        {
            string cellName = CellNameTextBox.Text.Trim();
            string cellStart = CellStartTextBox.Text.Trim();
            string cellEnd = CellEndTextBox.Text.Trim();
            if (string.IsNullOrEmpty(cellName) || string.IsNullOrEmpty(cellStart) || string.IsNullOrEmpty(cellEnd))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin.", "Thiếu thông tin", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            string rowInfo = $"Ô chứa tên: {cellName} | Ô bắt đầu: {cellStart} | Ô kết thúc: {cellEnd}";
            AddedRowsListBox.Items.Add(rowInfo);
            // Xóa nội dung các ô nhập
            CellNameTextBox.Text = "";
            CellStartTextBox.Text = "";
            CellEndTextBox.Text = "";
        }

        private void DeleteRowButton_Click(object sender, RoutedEventArgs e)
        {
            if (AddedRowsListBox.SelectedIndex >= 0)
            {
                AddedRowsListBox.Items.RemoveAt(AddedRowsListBox.SelectedIndex);
            }
            else
            {
                MessageBox.Show("Hãy chọn một dòng để xóa.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void SaveRowsToFile()
        {
            var lines = new List<string>();
            foreach (var item in AddedRowsListBox.Items)
            {
                lines.Add(item.ToString());
            }
            lines.Add("KEYSTRING:" + (KeyStringTextBox.Text ?? ""));
            System.IO.File.WriteAllLines(SaveFilePath, lines);
        }

        private void LoadRowsFromFile()
        {
            if (!System.IO.File.Exists(SaveFilePath)) return;
            var lines = System.IO.File.ReadAllLines(SaveFilePath);
            AddedRowsListBox.Items.Clear();
            foreach (var line in lines)
            {
                if (line.StartsWith("KEYSTRING:"))
                {
                    KeyStringTextBox.Text = line.Substring("KEYSTRING:".Length);
                }
                else
                {
                    AddedRowsListBox.Items.Add(line);
                }
            }
        }

        protected override void OnContentRendered(EventArgs e)
        {
            base.OnContentRendered(e);
            LoadRowsFromFile();
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            SaveRowsToFile();
            base.OnClosing(e);
        }

        private void ShowInfoButton_Click(object sender, RoutedEventArgs e)
        {
            DebugTextBox.Clear();
            ShowInfor();
        }

        // Chuyển A1 -> (row,col) (0-based)
        private (int row, int col) ParseCell(string cellAddress)
        {
            int col = 0;
            int i = 0;

            // Lấy phần chữ cái (cột)
            while (i < cellAddress.Length && char.IsLetter(cellAddress[i]))
            {
                col *= 26;
                col += (char.ToUpper(cellAddress[i]) - 'A' + 1);
                i++;
            }

            // Phần số là hàng
            string rowPart = cellAddress.Substring(i);
            int row = int.Parse(rowPart);

            return (row - 1, col - 1); // Excel: 1-based, NPOI: 0-based
        }
        public string GetCellValueFromSheet(string cellAddress, string filePath, int sheetIndex = 0)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !System.IO.File.Exists(filePath))
                return null;

            try
            {
                IWorkbook workbook;
                using (var fs = new System.IO.FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read))
                {
                    if (filePath.EndsWith(".xls"))
                        workbook = new HSSFWorkbook(fs);
                    else if (filePath.EndsWith(".xlsx"))
                        workbook = new XSSFWorkbook(fs);
                    else
                        throw new Exception("Định dạng file không hỗ trợ");

                    var sheet = workbook.GetSheetAt(sheetIndex);
                    var (row, col) = ParseCell(cellAddress); // dùng hàm ParseCell bạn đã có
                    var dataRow = sheet.GetRow(row);
                    if (dataRow == null) return null;
                    var dataCell = dataRow.GetCell(col);
                    return dataCell?.ToString();
                }
            }
            catch
            {
                return null;
            }
        }
        public async void ProcessCellRange(string start, string end, string filePath, int sheetIndex = 0)
        {
            var (startRow, startCol) = ParseCell(start.Trim());
            var (endRow, endCol) = ParseCell(end.Trim());

            using (var fs = new System.IO.FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read))
            {
                IWorkbook workbook;
                if (filePath.EndsWith(".xls"))
                    workbook = new HSSFWorkbook(fs);
                else if (filePath.EndsWith(".xlsx"))
                    workbook = new XSSFWorkbook(fs);
                else
                    throw new Exception("Định dạng file không hỗ trợ");

                var sheet = workbook.GetSheetAt(sheetIndex);

                AppendDebug("Xử lý thông tin...");
                for (int r = startRow; r <= endRow; r++)
                {
                    var row = sheet.GetRow(r);
                    if (row == null) continue;

                    for (int c = startCol; c <= endCol; c++)
                    {
                        var cell = row.GetCell(c);
                        string cellValue = cell?.ToString();
                        if (cellValue != null && !string.IsNullOrEmpty(cellValue))
                        {
                            AppendDebug(cellValue);
                        }
                        // hoặc xử lý gì đó với cellValue...
                    }
                }
            }
        }
        public List<WorkDay> ReadWorkDaysFromRange(string filePath, string startCell, string endCell, int sheetIndex = 0)
        {
            var workDays = new List<WorkDay>();

            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook;
                if (filePath.EndsWith(".xls"))
                    workbook = new HSSFWorkbook(fs);
                else
                    workbook = new XSSFWorkbook(fs);

                var sheet = workbook.GetSheetAt(sheetIndex);

                var (startRow, startCol) = ParseCell(startCell);
                var (endRow, endCol) = ParseCell(endCell);

                for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
                {
                    var row = sheet.GetRow(rowIndex);
                    if (row == null) continue;

                    var dayCell = row.GetCell(startCol);
                    string dayString = dayCell?.ToString()?.Trim();
                    if (string.IsNullOrEmpty(dayString)) continue;

                    var timeChecking = new List<string>();
                    for (int colIndex = startCol + 1; colIndex <= endCol; colIndex++)
                    {
                        var cell = row.GetCell(colIndex);
                        var value = cell?.ToString()?.Trim();
                        if (!string.IsNullOrEmpty(value))
                        {
                            timeChecking.Add(value);
                        }
                    }

                    workDays.Add(new WorkDay(dayString, timeChecking));
                }
            }

            return workDays;
        }
        public void ShowInfor()
        {
            string inputPath = InputFilePathTextBox.Text;
            if (string.IsNullOrWhiteSpace(inputPath) || !System.IO.File.Exists(inputPath))
            {
                AppendDebug("Vui lòng chọn file Excel input hợp lệ.");
                return;
            }
            string keyString = KeyStringTextBox.Text.Trim();
            if (string.IsNullOrEmpty(keyString))
            {
                AppendDebug("Vui lòng nhập Key String.");
                return;
            }
            try
            {
                IWorkbook workbook;
                using (var fs = new System.IO.FileStream(inputPath, System.IO.FileMode.Open, System.IO.FileAccess.Read))
                {
                    if (inputPath.EndsWith(".xls"))
                        workbook = new HSSFWorkbook(fs);
                    else if (inputPath.EndsWith(".xlsx"))
                        workbook = new XSSFWorkbook(fs);
                    else
                        throw new Exception("Định dạng file không hỗ trợ");

                    for (int sheetIdx = 0; sheetIdx < workbook.NumberOfSheets; sheetIdx++)
                    {
                        var sheet = workbook.GetSheetAt(sheetIdx);
                        bool hasKeyString = false;
                        for (int rowIdx = 0; rowIdx <= sheet.LastRowNum && !hasKeyString; rowIdx++)
                        {
                            var row = sheet.GetRow(rowIdx);
                            if (row == null) continue;
                            for (int colIdx = 0; colIdx < row.LastCellNum; colIdx++)
                            {
                                var cell = row.GetCell(colIdx);
                                if (cell != null && cell.CellType == CellType.String && cell.StringCellValue.Trim() == keyString)
                                {
                                    hasKeyString = true;
                                    break;
                                }
                            }
                        }
                        if (!hasKeyString) continue;
                        AppendDebug($"--- Sheet: {sheet.SheetName} chứa Key String ---");
                        foreach (var item in AddedRowsListBox.Items)
                        {
                            // Định dạng: Ô chứa tên: {cellName} | Ô bắt đầu: {cellStart} | Ô kết thúc: {cellEnd}
                            string line = item.ToString();
                            var parts = line.Split('|');
                            if (parts.Length < 3) continue;
                            string cellName = parts[0].Replace("Ô chứa tên:", "").Trim();
                            string cellStart = parts[1].Replace("Ô bắt đầu:", "").Trim();
                            string cellEnd = parts[2].Replace("Ô kết thúc:", "").Trim();
                            // Tìm tên nhân viên
                            string empName = GetCellValueFromSheet(cellName, inputPath, sheetIdx);
                            AppendDebug($"Tên nhân viên: {empName}");
                            // Tìm dải ô Time Card
                            List<WorkDay> workDays = ReadWorkDaysFromRange(inputPath, cellStart, cellEnd,  sheetIdx);
                            for (int i = 0; i < workDays.Count; i++)
                            {
                                AppendDebug("-------" + workDays[i].dayString);
                                for (int j = 0; j < workDays[i].timeChecking.Count; j++)
                                {
                                    AppendDebug(workDays[i].timeChecking[j]);
                                }

                            }
                            /*List<string> thongtin = ProcessCellRange(cellStart, cellEnd, inputPath, sheetIdx);
                            if (thongtin.Count > 0)
                            {
                                AppendDebug($"Dải ô Time Card: {cellStart} - {cellEnd}");
                                AppendDebug("Thông tin Time Card:");
                                    await Task.Delay(1);
                                foreach (var info in thongtin)
                                {
                                    AppendDebug(info);
                                    await Task.Delay(1);
                                }
                            }*/
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                AppendDebug($"Lỗi khi đọc file: {ex.Message}");
            }
        }
        
    }
    [System.Serializable]
    public class WorkDay
    {
        public string dayString;
        public List<string> timeChecking;

        public WorkDay(string dayString, List<string> timeChecking)
        {
            this.dayString = dayString;
            this.timeChecking = timeChecking;
        }
    }
}
