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
using System.Threading;
using MahApps.Metro.Controls;

namespace CountWodayWpf
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
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
            //ExportSample();
            return;
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
        private void AddRowButton_ClickFee(object sender, RoutedEventArgs e)
        {
            string cellFee = FeeTextBox.Text.Trim();
            if (string.IsNullOrEmpty(cellFee) || int.TryParse(cellFee, out _))
            {
                MessageBox.Show("Vui lòng nhập số tiền phạt đúng dạng số", "Thiếu thông tin", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            string rowInfo = $"{cellFee}";
            AddedRowsFeeListBox.Items.Add(rowInfo);
            // Xóa nội dung các ô nhập
            FeeTextBox.Text = "";
        }

        private void DeleteRowButton_ClickFee(object sender, RoutedEventArgs e)
        {
            if (AddedRowsFeeListBox.SelectedIndex >= 0)
            {
                AddedRowsFeeListBox.Items.RemoveAt(AddedRowsFeeListBox.SelectedIndex);
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
            foreach (var item in AddedRowsFeeListBox.Items)
            {
                lines.Add("Fee:" + item.ToString());
            }
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
                else if (line.StartsWith("Fee:"))
                {
                    AddedRowsFeeListBox.Items.Add(line.Substring("Fee:".Length));
                }
                else
                {
                    AddedRowsListBox.Items.Add(line);
                }
            }
            LoadFee();
        }

        protected override void OnContentRendered(EventArgs e)
        {
            base.OnContentRendered(e);
            LoadRowsFromFile();
        }
        public List<int> feeList;
        public void LoadFee()
        {
            feeList = new List<int>();
            for (int i = 0; i < AddedRowsFeeListBox.Items.Count; i++)
            {
                int fee;
                if(int.TryParse(AddedRowsFeeListBox.Items[i].ToString(), out fee))
                {
                    feeList.Add(fee);
                }
                else
                {
                    AppendDebug($"{AddedRowsFeeListBox.Items[i].ToString()} Định dạng tiền phạt đang bị sai, hãy sửa lại để tiền phạt hoạt động đúng.");
                }
            }
        }
        public int GetFee(int penaltyCount)
        {
            if (feeList == null || feeList.Count == 0 || penaltyCount <= 0)
                return 0;

            int total = 0;

            for (int i = 0; i < penaltyCount; i++)
            {
                if (i < feeList.Count)
                    total += feeList[i];
                else
                    total += feeList[feeList.Count - 1]; // dùng mức phạt cuối
            }

            return total;
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
                            if (cell != null && cell.CellType == CellType.Numeric && DateUtil.IsCellDateFormatted(cell))
                            {
                                timeChecking.Add(GetRoundedTimeString(cell.DateCellValue)); // Hoặc định dạng khác nếu cần
                            }
                            else
                            {
                                timeChecking.Add(value);
                            }
                        }
                    }

                    workDays.Add(new WorkDay(dayString, timeChecking));
                }
            }

            return workDays;
        }
        public static string GetRoundedTimeString(DateTime dt)
        {
            // Nếu giây >= 30 thì cộng thêm 1 phút
            if (dt.Second >= 30)
                dt = dt.AddMinutes(1);

            // Trả về dạng HH:mm, bỏ giây
            return new DateTime(dt.Year, dt.Month, dt.Day, dt.Hour, dt.Minute, 0).ToString("HH:mm");
        }
        BoxWriter writer;
        public async void ShowInfor()
        {
            LoadingText.Visibility = Visibility.Visible;
            ProgressBar.Visibility = Visibility.Visible;
            string inputPath = InputFilePathTextBox.Text;
            SetupSheet("Kết quả chấm công", new int[] { 18, 30 });
            writer = new BoxWriter(sheet, borderedStyle, 1);
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
                        foreach (var item in AddedRowsListBox.Items)
                        {
                            await Task.Run(() =>
                            {
                                // Gọi hàm xử lý nặng ở đây
                                Thread.Sleep(1); // giả lập xử lý
                                                    // hoặc ProcessData(); 
                            });
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
                            HandleDayWork(ReadWorkDaysFromRange(inputPath, cellStart, cellEnd, sheetIdx), empName);

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
                    SaveWorkbook(writeWorkbook);
                }
            }
            catch (Exception ex)
            {
                AppendDebug($"Lỗi khi đọc file: {ex.Message}");
            }
            LoadingText.Visibility = Visibility.Collapsed;
            ProgressBar.Visibility = Visibility.Collapsed;
        }
        TimeSpan h11 = new TimeSpan(11, 0, 0);
        TimeSpan h9 = new TimeSpan(9, 30, 0);
        TimeSpan h2 = new TimeSpan(2, 0, 0);
        TimeSpan h3 = new TimeSpan(3, 0, 0);
        TimeSpan h6 = new TimeSpan(6, 0, 0);
        TimeSpan h17 = new TimeSpan(17, 0, 0);
        public void HandleDayWork(List<WorkDay> _workDays,string nameEmp)
        {
            if (!TimeSpan.TryParse(StartTimeTextBox.Text, out TimeSpan allowedStartTime))
            {
                MessageBox.Show("Giờ vào làm không hợp lệ. Định dạng đúng: HH:mm", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            int feeSum = 0;
            int lateDayCount = 0;
            string LateDays = "";
            int NoCheckinCount = 0;
            string NoCheckin = "";
            int NoCheckoutCount = 0;
            string NoCheckout = "";
            int HalfDayCount = 0;
            string HalfDays = "";
            float dayWork = 0;

            for (int i = 0; i < _workDays.Count; i++)
            {
                var workDay = _workDays[i];

                if (workDay.timeChecking == null || workDay.timeChecking.Count == 0)
                    continue;


                if (TimeSpan.TryParse(_workDays[i].timeChecking[0], out TimeSpan parsedTime))
                {
                    TimeSpan parsedTimeEnd = TimeSpan.Parse(_workDays[i].timeChecking[_workDays[i].timeChecking.Count - 1]);
                    var duration = parsedTimeEnd - parsedTime;
                    //Đếm ngày đi muộn 
                    if (parsedTime > allowedStartTime && parsedTime < h11)
                    {
                        lateDayCount++;
                        LateDays += $"{i + 1}, "; // +1 để đánh số ngày bắt đầu từ 1
                    }

                    //Đếm ngày không checkin
                    if (parsedTime > h9 && duration < h2)
                    {
                        NoCheckinCount++;
                        NoCheckin += $"{i + 1}, "; // +1 để đánh số ngày bắt đầu từ 1
                    }

                    //Đếm ngày không checkout
                    if (parsedTime < h17 && duration < h2)
                    {
                        NoCheckoutCount++;
                        NoCheckout += $"{i + 1}, "; // +1 để đánh số ngày bắt đầu từ 1
                    }

                    //Đếm ngày không checkout
                    if (duration > h3 && duration < h6)
                    {
                        HalfDayCount++;
                        HalfDays += $"{i + 1}, "; // +1 để đánh số ngày bắt đầu từ 1
                        dayWork += 0.5f;
                    }
                    else
                    {
                        dayWork += 1;
                    }

                }
            }

            if (lateDayCount > 0 && LateDays.Length >= 2)
            {
                LateDays = LateDays.Substring(0, LateDays.Length - 2);
                feeSum += GetFee(lateDayCount);
            }
            if (NoCheckinCount > 0 && NoCheckin.Length >= 2)
            {
                NoCheckin = NoCheckin.Substring(0, NoCheckin.Length - 2);
                feeSum += GetFee(NoCheckinCount);
            }
            if (NoCheckoutCount > 0 && NoCheckout.Length >= 2)
            {
                NoCheckout = NoCheckout.Substring(0, NoCheckout.Length - 2);
                feeSum += GetFee(NoCheckoutCount);
            }
            if (HalfDayCount > 0 && HalfDays.Length >= 2)
                HalfDays = HalfDays.Substring(0, HalfDays.Length - 2);
            if (dayWork > 0)
            {
                AppendDebug("----------");
                AppendDebug($"Số ngày công: {dayWork}");
                AppendDebug("----------");
                AppendDebug($"Số hôm làm nửa ngày: {HalfDayCount}");
                AppendDebug($"Vào ngày: {HalfDays}");
                AppendDebug("----------");
                AppendDebug($"Số hôm đi muộn: {lateDayCount}");
                AppendDebug($"Vào ngày: {LateDays}");
                AppendDebug("----------");
                AppendDebug($"Số hôm không check in: {NoCheckinCount}");
                AppendDebug($"Vào ngày: {NoCheckin}");
                AppendDebug("----------");
                AppendDebug($"Số hôm không check out: {NoCheckoutCount}");
                AppendDebug($"Vào ngày: {NoCheckout}");
                feeSum *= 1000;
                writer.WriteBox(new string[,]
                               {

                                    { "Tên", nameEmp },
                                    { "Số ngày công:", dayWork.ToString()},
                                    { "Làm nửa ngày:", HalfDayCount.ToString()},
                                    { "Vào ngày", HalfDays },
                                    { "Đi muộn:", lateDayCount.ToString()},
                                    { "Vào ngày", LateDays },
                                    { "Không check in:", NoCheckinCount.ToString()},
                                    { "Vào ngày", NoCheckin },
                                    { "Không check out:", NoCheckoutCount.ToString()},
                                    { "Vào ngày", NoCheckout },
                                    { "Tổng tiền phạt", feeSum.ToString()},
                               },

                                new bool[,] {
                                    { true, true },
                                    { true, false },
                                    { true, false },
                                    { true, false },
                                    { true, false },
                                    { true, false },
                                    { true, false },
                                    { true, false },
                                    { true, false },
                                    { true, false },
                                    { true, false },
                                },
                                new string[,] {
                                    {"LightYellow", "" },
                                    { "LightGreen", "LightGreen" },
                                    { "LightGreen", "LightGreen" },
                                    { "", "" },
                                    { "LightGreen", "LightGreen" },
                                    { "", "" },
                                    { "LightGreen", "LightGreen" },
                                    { "", "" },
                                    { "LightGreen", "LightGreen" },
                                    { "", "" },
                                    { "LightBlue", "LightBlue" },
                                }

                               );
            }
            else
            {
                AppendDebug($"Không đi làm");
                writer.WriteBox(new string[,]
                               {
                                    { "Tên", nameEmp },
                                    { "Không đi làm", " "},
                               },
                               new bool[,] {
                                    { true, true },
                                    { false, false },
                                },
                                new string[,] {
                                    {"LightYellow", "" },
                                    { "", "" },
                                });
            }
            AppendDebug("----------");
            AppendDebug("      ");
            writer.WriteBox(new string[,]
                               {
                                    { "", "" }
                               });


        }
        public IWorkbook writeWorkbook;
        public ISheet sheet;
        public ICellStyle borderedStyle;
        public void SetupSheet(string sheetName, int[] columnWidths)
        {
            writeWorkbook = new XSSFWorkbook();
            sheet = writeWorkbook.CreateSheet(sheetName);

            // Set column widths (width in characters * 256)
            for (int i = 0; i < columnWidths.Length; i++)
            {
                sheet.SetColumnWidth(i, columnWidths[i] * 256);
            }

            // Tạo style có border
            borderedStyle = writeWorkbook.CreateCellStyle();
            borderedStyle.BorderTop = BorderStyle.Thin;
            borderedStyle.BorderBottom = BorderStyle.Thin;
            borderedStyle.BorderLeft = BorderStyle.Thin;
            borderedStyle.BorderRight = BorderStyle.Thin;
        }
        public void SaveWorkbook(IWorkbook workbook)
        {
            using (var fs = new FileStream(OutputFilePathTextBox.Text, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
                AppendDebug("EXPORT FILE THÀNH CÔNG");
                 var result = MessageBox.Show("EXPORT FILE THÀNH CÔNG!\nBạn có muốn mở file ngay bây giờ?",
                              "Thành công",
                              MessageBoxButton.YesNo,
                              MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    System.Diagnostics.Process.Start("explorer.exe", OutputFilePathTextBox.Text);
                }
            }
        }

        public void ExportSample()
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("MySheet");

            IRow row = sheet.CreateRow(0);
            row.CreateCell(0).SetCellValue("Hello");
            row.CreateCell(1).SetCellValue("World");


            using (var fs = new FileStream(OutputFilePathTextBox.Text, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
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

    public class BoxWriter
    {
        private readonly ISheet _sheet;
        private readonly IWorkbook _workbook;
        private readonly ICellStyle _baseStyle;
        private int _currentRow;

        private readonly Dictionary<string, ICellStyle> _styleCache = new();

        // Map tên màu sang chỉ số màu tương thích với NPOI 2.6+
        private readonly Dictionary<string, short> _colorMap = new()
        {
            { "Red", IndexedColors.Red.Index },
            { "Green", IndexedColors.Green.Index },
            { "Yellow", IndexedColors.Yellow.Index },
            { "LightGreen", IndexedColors.LightGreen.Index },
            { "LightBlue", IndexedColors.LightCornflowerBlue.Index },
            { "LightYellow", IndexedColors.LightOrange.Index },
            { "Grey25", IndexedColors.Grey25Percent.Index },
            { "Orange", IndexedColors.Orange.Index },
            { "None", -1 }
        };

        public BoxWriter(ISheet sheet, ICellStyle baseStyle, int startRow = 0)
        {
            _sheet = sheet;
            _workbook = sheet.Workbook;
            _baseStyle = baseStyle;
            _currentRow = startRow;
        }

        private ICellStyle GetOrCreateStyle(bool isBold, string colorName)
        {
            string key = $"{isBold}_{colorName ?? "None"}";

            if (_styleCache.TryGetValue(key, out var cachedStyle))
                return cachedStyle;

            var font = _workbook.CreateFont();
            font.IsBold = isBold;

            var style = _workbook.CreateCellStyle();
            style.CloneStyleFrom(_baseStyle);
            style.SetFont(font);

            if (!string.IsNullOrWhiteSpace(colorName) &&
                _colorMap.TryGetValue(colorName, out var colorIndex) &&
                colorIndex >= 0)
            {
                style.FillForegroundColor = colorIndex;
                style.FillPattern = FillPattern.SolidForeground;
            }

            _styleCache[key] = style;
            return style;
        }

        public void WriteBox(
            string[,] content,
            bool[,] isBold = null,
            string[,] fillColorNames = null,
            int startCol = 0,
            int rowSpacing = 1)
        {
            int rowCount = content.GetLength(0);
            int colCount = content.GetLength(1);

            for (int r = 0; r < rowCount; r++)
            {
                IRow row = _sheet.GetRow(_currentRow + r) ?? _sheet.CreateRow(_currentRow + r);

                for (int c = 0; c < colCount; c++)
                {
                    string text = content[r, c] ?? "";
                    ICell cell = row.GetCell(startCol + c) ?? row.CreateCell(startCol + c);
                    cell.SetCellValue(text);

                    bool bold = isBold != null &&
                                isBold.GetLength(0) > r &&
                                isBold.GetLength(1) > c &&
                                isBold[r, c];

                    string colorName = fillColorNames != null &&
                                       fillColorNames.GetLength(0) > r &&
                                       fillColorNames.GetLength(1) > c
                                       ? fillColorNames[r, c]
                                       : null;

                    cell.CellStyle = GetOrCreateStyle(bold, colorName);
                }
            }

            _currentRow += rowCount + rowSpacing;
        }

        public int GetCurrentRow() => _currentRow;
    }

}