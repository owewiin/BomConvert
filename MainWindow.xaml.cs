using ClosedXML.Excel;
using Compo;
using DocumentFormat.OpenXml.Bibliography;
using ExcelDataReader;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace SchBom_Convert
{
    public partial class MainWindow : Window
    {
        public List<string> SelectedFilePaths { get; set; } = new List<string>();
        public Dictionary<string, List<BomPreviewItem>> FileDataDict { get; set; } = new Dictionary<string, List<BomPreviewItem>>();
        // 頁面類型枚舉
        public enum PanelType
        {
            Main,
            Settings,
            Preview,
            ConversionInfo,
            VersionHistory
        }
        public MainWindow()
        {
            InitializeComponent();
            LoadAutoOpenSetting(); // 載入之前儲存的設定
            LoadStaffSettings(); // 載入人員設定
            ResetCustomerProductFields(true); // 載入客戶與產品名稱設定
            this.DataContext = this;
            CurrentDate = DateTime.Now.ToString("yyyy/MM/dd tt HH 時");
            this.PreviewKeyDown += MainWindow_PreviewKeyDown;
            LogUsage("啟動程式");
        }

        private void MainWindow_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (Keyboard.IsKeyDown(Key.LeftCtrl) && Keyboard.IsKeyDown(Key.LeftAlt) && e.Key == Key.L)
            {
                var adminWin = new AdminLogWindow();
                adminWin.Show();
            }
        }
        private void VersionHistoryButton_Click(object sender, RoutedEventArgs e)
        {
            LoadVersionHistory();
            ShowPanel(PanelType.VersionHistory);
        }

        // 儲存自動開啟檔案的設定
        private bool isAutoOpenFileEnabled = false;
        public string CurrentDate { get; set; }
        public ObservableCollection<BomItem> BomList { get; set; } = new();
        public ObservableCollection<BomPreviewItem> BomPreviewList { get; set; } = new();
        public string SelectedRDAssistant { get; set; } = "Fish";
        public string SelectedLayoutPerson { get; set; } = "WEI";
        public string SelectedCircuitDesigner { get; set; } = "LSP";
        
        // 儲存各類人員的名單列表
        public ObservableCollection<string> RDAssistantList { get; set; } = new ObservableCollection<string> { "Peggy", "Fish" };
        public ObservableCollection<string> LayoutPersonList { get; set; } = new ObservableCollection<string> { "未定", "WEI", "Jane", "Wuct", "JDLee", "Jason" };
        public ObservableCollection<string> CircuitDesignerList { get; set; } = new ObservableCollection<string> { "未定", "LSP", "Jane", "Jason", "Kevin", "Yanchi" };
        public ObservableCollection<string> ConversionMessages { get; set; } = new();

        // 在現有屬性區域新增這兩行
        public string CustomerName { get; set; } = "";
        public string ProductName { get; set; } = "";

        public class BomItem
        {
            public string? ItemNumber { get; set; }
            public string? ShortName { get; set; }
            public string? Spec { get; set; }
            public string? Package { get; set; }
            public string? Quantity { get; set; }
            public string? PartNumber { get; set; }
            public string? UnitPrice { get; set; }
            public string? Brand { get; set; }
            public string? Note { get; set; }
            public string? Process { get; set; }
            public string? AltPart1 { get; set; }
            public string? ChinaPart { get; set; }
            public string? ChinaBrand { get; set; }
            public string? ChinaAlt { get; set; }
            public string? AltPart2 { get; set; }
        }
        public class BomPreviewItem
        {
            public int? Index { get; set; }
            public string? PartName { get; set; }
            public string? Spec { get; set; }
            public string? Package { get; set; }
            public string? Code { get; set; }
            public int Quantity { get; set; }
            public decimal UnitPrice { get; set; }
            public string? Vendor { get; set; }
            public string? VendorCN { get; set; }
            public string? Note { get; set; }
            public bool IsAltLine { get; set; }
            public string? Alt1 { get; set; }
            public string? Alt2 { get; set; }
            public string? AltCN { get; set; }
            public decimal Subtotal => UnitPrice * Quantity;
            public string QuantityDisplay => IsAltLine ? "" : Quantity.ToString();
            public string UnitPriceDisplay => IsAltLine ? "" : UnitPrice.ToString("0.##");
            public string SubtotalDisplay => IsAltLine ? "" : Subtotal.ToString("0.##");
            public string IndexDisplay => Index?.ToString() ?? "";
            public string Category
            {
                get
                {
                    if (IsAltLine) return "";
                    if (PartName?.Contains("SMD", StringComparison.OrdinalIgnoreCase) == true)
                        return "SMD 零件";
                    if (PartName?.Contains("DIP", StringComparison.OrdinalIgnoreCase) == true ||
                        PartName?.Contains("PTH", StringComparison.OrdinalIgnoreCase) == true ||
                        PartName?.Contains("THT", StringComparison.OrdinalIgnoreCase) == true)
                        return "DIP 零件";
                    return "其他 零件";
                }
            }
        }
        // 選擇多檔案按鈕事件
        private void ChooseMultipleFiles_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var openFileDialog = new OpenFileDialog
                {
                    Title = "選擇多個檔案進行合併",
                    Filter = "Excel 檔案 (*.xlsx;*.xls;*.xlt)|*.xlsx;*.xls;*.xlt|所有檔案 (*.*)|*.*",
                    Multiselect = true
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    SelectedFilePaths = openFileDialog.FileNames.ToList();
                    UpdateSelectedFilesList();

                    // 清空原有單檔資料
                    FilePathTextBox.Text = "";
                    BomPreviewList.Clear();

                    // 顯示合併匯出按鈕
                    MergeExportButton.Visibility = Visibility.Visible;
                    //MergeExportPreviewButton.Visibility = Visibility.Visible;  // 拿掉預覽部分的合併匯出按鈕

                    // 預先載入所有檔案資料
                    LoadMultipleFilesData();

                    AddConversionMessage($"[INFO] 已選擇 {SelectedFilePaths.Count} 個檔案進行合併");
                    MessageBox.Show($"已選擇 {SelectedFilePaths.Count} 個檔案\n點擊「合併匯出」開始處理",
                                    "選擇完成", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"選擇多檔案時發生錯誤：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
                AddConversionMessage($"[ERROR] ChooseMultipleFiles_Click 發生錯誤: {ex}");
            }
        }

        // 更新已選檔案列表顯示
        private void UpdateSelectedFilesList()
        {
            SelectedFilesPanel.Children.Clear();

            if (SelectedFilePaths.Count > 0)
            {
                SelectedFilesScrollViewer.Visibility = Visibility.Visible;

                for (int i = 0; i < SelectedFilePaths.Count; i++)
                {
                    var fileInfo = new StackPanel
                    {
                        Orientation = Orientation.Horizontal,
                        Margin = new Thickness(0, 2, 0, 2)
                    };

                    var fileText = new TextBlock
                    {
                        Text = $"{i + 1}. {Path.GetFileName(SelectedFilePaths[i])}",
                        FontSize = 12,
                        Foreground = Brushes.DarkBlue,
                        Margin = new Thickness(5, 0, 10, 0)
                    };

                    var removeButton = new Button
                    {
                        Content = "✖",
                        Width = 20,
                        Height = 20,
                        FontSize = 10,
                        Background = Brushes.LightCoral,
                        Foreground = Brushes.White,
                        Tag = i
                    };
                    removeButton.Click += RemoveSelectedFile_Click;

                    fileInfo.Children.Add(fileText);
                    fileInfo.Children.Add(removeButton);
                    SelectedFilesPanel.Children.Add(fileInfo);
                }
            }
            else
            {
                SelectedFilesScrollViewer.Visibility = Visibility.Collapsed;
                MergeExportButton.Visibility = Visibility.Collapsed;
                //MergeExportPreviewButton.Visibility = Visibility.Collapsed;
            }
        }

        // 移除選擇的檔案
        private void RemoveSelectedFile_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is Button button && button.Tag is int index)
                {
                    string fileName = Path.GetFileName(SelectedFilePaths[index]);
                    SelectedFilePaths.RemoveAt(index);

                    // 從資料字典中移除
                    var keyToRemove = FileDataDict.Keys.FirstOrDefault(k => Path.GetFileName(k) == fileName);
                    if (keyToRemove != null)
                        FileDataDict.Remove(keyToRemove);

                    UpdateSelectedFilesList();
                    AddConversionMessage($"[INFO] 已移除檔案：{fileName}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"移除檔案時發生錯誤：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
                AddConversionMessage($"[ERROR] RemoveSelectedFile_Click 發生錯誤: {ex}");
            }
        }

        // 載入多檔案資料
        private void LoadMultipleFilesData()
        {
            try
            {
                FileDataDict.Clear();
                int successCount = 0;
                int errorCount = 0;

                foreach (string filePath in SelectedFilePaths)
                {
                    try
                    {
                        var tempBomList = new List<BomPreviewItem>();
                        if (ProcessSingleFileForMerge(filePath, tempBomList))
                        {
                            FileDataDict[filePath] = tempBomList;
                            successCount++;
                            AddConversionMessage($"[INFO] 成功載入：{Path.GetFileName(filePath)} ({tempBomList.Count(x => !x.IsAltLine)} 項零件)");
                        }
                        else
                        {
                            errorCount++;
                            AddConversionMessage($"[ERROR] 載入失敗：{Path.GetFileName(filePath)}");
                        }
                    }
                    catch (Exception fileEx)
                    {
                        errorCount++;
                        AddConversionMessage($"[ERROR] 載入檔案 {Path.GetFileName(filePath)} 發生錯誤: {fileEx.Message}");
                    }
                }

                AddConversionMessage($"[INFO] 多檔案載入完成 - 成功：{successCount} 失敗：{errorCount}");

                if (errorCount > 0)
                {
                    MessageBox.Show($"載入完成！\n成功：{successCount} 個檔案\n失敗：{errorCount} 個檔案\n詳細資訊請查看轉換資訊頁面",
                                    "載入結果", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"載入多檔案資料時發生錯誤：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
                AddConversionMessage($"[ERROR] LoadMultipleFilesData 發生錯誤: {ex}");
            }
        }

        // 處理單個檔案用於合併
        private bool ProcessSingleFileForMerge(string fullPath, List<BomPreviewItem> bomList)
        {
            try
            {
                string extension = Path.GetExtension(fullPath).ToLower();
                if (extension != ".xls" && extension != ".xlsx" && extension != ".xlt")
                    return false;

                if (!File.Exists(fullPath))
                    return false;

                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                using var stream = File.Open(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
                using DataSet dataSet = reader.AsDataSet();

                if (dataSet?.Tables?.Count == 0 || dataSet.Tables[0].Rows.Count <= 5)
                    return false;

                var table = dataSet.Tables[0];
                var rawMainList = new List<BomPreviewItem>();

                for (int i = 5; i < table.Rows.Count; i++)
                {
                    try
                    {
                        var row = table.Rows[i];
                        if (row.ItemArray.All(cell => string.IsNullOrWhiteSpace(cell?.ToString())))
                            continue;

                        string partNumber = GetSafeStringValue(row, 5);
                        if (string.IsNullOrWhiteSpace(partNumber))
                            continue;

                        var parts = partNumber.Split(new[] { ',', '，', ';', '；' }, StringSplitOptions.RemoveEmptyEntries);
                        int qty = parts.Length;
                        if (qty == 0) continue;

                        decimal price = 0;
                        string priceString = GetSafeStringValue(row, 6);
                        if (!string.IsNullOrWhiteSpace(priceString))
                            decimal.TryParse(priceString, out price);

                        string spec = NormalizeNCFormat(GetSafeStringValue(row, 2));

                        rawMainList.Add(new BomPreviewItem
                        {
                            IsAltLine = false,
                            PartName = GetFinalPartName(GetSafeStringValue(row, 1, "零件名稱"), partNumber),
                            Spec = spec,
                            Package = GetSafeStringValue(row, 3, "包裝"),
                            Quantity = qty,
                            Code = partNumber,
                            UnitPrice = price,
                            Vendor = GetSafeStringValue(row, 7),
                            VendorCN = string.IsNullOrWhiteSpace(GetSafeStringValue(row, 13))
                            ? GetSafeStringValue(row, 12) : GetSafeStringValue(row, 13),
                            Note = GetSafeStringValue(row, 8),
                            Alt1 = GetSafeStringValue(row, 10),
                            AltCN = GetSafeStringValue(row, 13),
                            Alt2 = GetSafeStringValue(row, 14)
                        });
                    }
                    catch
                    {
                        continue;
                    }
                }

                if (rawMainList.Count == 0)
                    return false;

                // 排序和合併邏輯（與原有邏輯相同）
                var sortedList = rawMainList.OrderBy(GetMainSortKey).ToList();

                var mergedList = MergeDuplicateItems(sortedList);
                int index = 1;
                var grouped = mergedList
                    .OrderBy(GetCategoryGroupOrder)
                    .GroupBy(p => p.Category);

                foreach (var group in grouped)
                {
                    bomList.Add(new BomPreviewItem
                    {
                        IsAltLine = true,
                        PartName = group.Key.Replace(" 零件", "") + "料表",
                        Spec = "",
                        Vendor = "",
                        VendorCN = "",
                        Note = ""
                    });

                    foreach (var main in group)
                    {
                        main.Index = index++;
                        bomList.Add(main);

                        void AddAltRow(string? text, string label)
                        {
                            if (string.IsNullOrWhiteSpace(text)) return;
                            try
                            {
                                var (spec, vendor, flagged) = AltPartRules.Parse(text.Trim());
                                bomList.Add(new BomPreviewItem
                                {
                                    IsAltLine = true,
                                    PartName = $"{label}:",
                                    Spec = spec,
                                    VendorCN = vendor,
                                    Index = null
                                });
                            }
                            catch
                            {
                                bomList.Add(new BomPreviewItem
                                {
                                    IsAltLine = true,
                                    PartName = $"{label}:",
                                    Spec = text,
                                    VendorCN = "",
                                    Index = null
                                });
                            }
                        }

                        AddAltRow(main.Alt1, "替代料");
                        AddAltRow(main.Alt2, "替代料2");
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        // 合併匯出按鈕事件
        private void MergeExportButton_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedFilePaths.Count == 0)
            {
                MessageBox.Show("請先選擇要合併的檔案", "沒有選擇檔案", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (FileDataDict.Count == 0)
            {
                MessageBox.Show("沒有有效的資料可以匯出，請重新載入檔案", "沒有資料", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            CustomerName = CustomerNameTextBox.Text?.Trim() ?? "";
            ProductName = ProductNameTextBox.Text?.Trim() ?? "";

            var saveFileDialog = new SaveFileDialog
            {
                Title = "匯出合併的 BOM 資料",
                Filter = "Excel 檔案 (*.xlsx)|*.xlsx",
                DefaultExt = "xlsx",
                FileName = $"合併BOM_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    ExportMergedToExcel(saveFileDialog.FileName);
                    MessageBox.Show($"合併匯出成功！\n已處理 {FileDataDict.Count} 個檔案\n檔案已儲存至：{saveFileDialog.FileName}",
                                  "匯出完成", MessageBoxButton.OK, MessageBoxImage.Information);
                    LogUsage("合併匯出Excel", Path.GetFileName(saveFileDialog.FileName));

                    if (isAutoOpenFileEnabled)
                        OpenFile(saveFileDialog.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"合併匯出失敗：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
                    AddConversionMessage($"[ERROR] 合併匯出失敗: {ex}");
                }
            }
        }

        // 匯出合併的Excel檔案
        private void ExportMergedToExcel(string filePath)
        {
            using var workbook = new XLWorkbook();

            foreach (var fileData in FileDataDict)
            {
                string originalFileName = Path.GetFileNameWithoutExtension(fileData.Key);
                string worksheetName = SanitizeWorksheetName(originalFileName);

                // 確保工作表名稱唯一
                int counter = 1;
                string uniqueWorksheetName = worksheetName;
                while (workbook.Worksheets.Any(ws => ws.Name == uniqueWorksheetName))
                {
                    uniqueWorksheetName = $"{worksheetName}_{counter}";
                    counter++;
                }

                var worksheet = workbook.Worksheets.Add(uniqueWorksheetName);
                var bomData = fileData.Value;

                SetColumnWidths(worksheet);
                ExportSingleWorksheet(worksheet, bomData, originalFileName);

                AddConversionMessage($"[INFO] 已匯出工作表：{uniqueWorksheetName} ({bomData.Count(x => !x.IsAltLine)} 項零件)");
            }

            workbook.SaveAs(filePath);
        }

        // 匯出單個工作表
        private void ExportSingleWorksheet(IXLWorksheet worksheet, List<BomPreviewItem> bomList, string originalFileName)
        {
            int currentRow = 5;
            var categories = bomList.Where(item => item.IsAltLine && item.PartName?.EndsWith("料表") == true).ToList();
            bool isFirstCategory = true;

            for (int categoryIndex = 0; categoryIndex < categories.Count; categoryIndex++)
            {
                var categoryItem = categories[categoryIndex];
                string categoryName = categoryItem.PartName ?? "";

                SetCategoryHeader(worksheet, currentRow, originalFileName, categoryName, ref isFirstCategory);
                SetTableHeaders(worksheet, currentRow);
                currentRow++;

                currentRow = ProcessCategoryItemsForWorksheet(worksheet, bomList, categoryName, currentRow, out decimal categoryTotal);
                SetCategorySubtotal(worksheet, currentRow, categoryName, categoryTotal);
                currentRow++;

                AddStaffInfoToCategory(worksheet, currentRow);
                currentRow += 5;

                if (categoryIndex < categories.Count - 1)
                {
                    AddEmptyRows(worksheet, currentRow, 2, 35, false);
                    currentRow += 2;
                }
            }

            SetGrandTotalForWorksheet(worksheet, currentRow - 3, bomList);
        }

        // 處理單個工作表的分類項目
        private int ProcessCategoryItemsForWorksheet(IXLWorksheet worksheet, List<BomPreviewItem> bomList, string categoryName, int currentRow, out decimal categoryTotal)
        {
            var (categoryItems, ncItems) = GetCategoryItemsFromList(bomList, categoryName);
            categoryTotal = 0;
            int itemNumber = 1;

            foreach (var item in categoryItems)
                currentRow = ProcessItemWithNumber(worksheet, item, currentRow, ref categoryTotal, ref itemNumber);

            if (ncItems.Count > 0)
            {
                AddEmptyRows(worksheet, currentRow, 1, 35, true);
                currentRow++;

                foreach (var ncItem in ncItems)
                    currentRow = ProcessNCItemWithNumber(worksheet, ncItem, currentRow, ref categoryTotal, ref itemNumber);
            }

            AddEmptyRows(worksheet, currentRow, 2, 35, true);
            return currentRow + 2;
        }

        // 從指定列表取得分類項目
        private (List<BomPreviewItem> categoryItems, List<BomPreviewItem> ncItems) GetCategoryItemsFromList(List<BomPreviewItem> bomList, string categoryName)
        {
            var categoryItems = new List<BomPreviewItem>();
            var ncItems = new List<BomPreviewItem>();
            bool foundCategory = false;

            foreach (var item in bomList)
            {
                if (item.IsAltLine && item.PartName == categoryName)
                {
                    foundCategory = true;
                    continue;
                }

                if (foundCategory)
                {
                    if (item.IsAltLine && item.PartName?.EndsWith("料表") == true)
                        break;

                    if (item.IsAltLine)
                        continue;

                    if (!item.IsAltLine)
                    {
                        if (IsNCItem(item.Spec))
                            ncItems.Add(item);
                        else
                            categoryItems.Add(item);
                    }
                }
            }

            return (categoryItems, ncItems);
        }

        // 設定單個工作表的總計
        private void SetGrandTotalForWorksheet(IXLWorksheet worksheet, int currentRow, List<BomPreviewItem> bomList)
        {
            decimal grandTotal = bomList.Where(item => !item.IsAltLine).Sum(item => item.Subtotal);

            var labelCell = worksheet.Cell(currentRow, 7);
            labelCell.Value = "總計：";
            labelCell.Style.Font.Bold = true;
            labelCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

            var totalCell = worksheet.Cell(currentRow, 8);
            totalCell.Value = grandTotal.ToString("0.##");
            var style = totalCell.Style;
            style.Font.Bold = true;
            style.Fill.BackgroundColor = XLColor.Yellow;
            style.Border.OutsideBorder = XLBorderStyleValues.Thick;

            worksheet.Row(currentRow).Height = 24.8;
        }
        private void CustomerNameTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            CustomerName = CustomerNameTextBox.Text?.Trim() ?? "";
            SaveCustomerProductSettings();
        }

        private void ProductNameTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            ProductName = ProductNameTextBox.Text?.Trim() ?? "";
            SaveCustomerProductSettings();
        }
        private void ResetCustomerProductFields(bool alsoPersist = true)
        {
            CustomerName = "";
            ProductName = "";

            if (CustomerNameTextBox != null) CustomerNameTextBox.Text = "";
            if (ProductNameTextBox != null) ProductNameTextBox.Text = "";

            if (alsoPersist) SaveCustomerProductSettings(); // 把「空值」寫回去，確保下次啟動也空白
        }
        // 載入版本歷史
        private void LoadVersionHistory()
        {
            try
            {
                string readmePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "README.md");

                if (File.Exists(readmePath))
                {
                    string readmeContent = File.ReadAllText(readmePath, Encoding.UTF8);
                    // 找版本歷史
                    int startIndex = readmeContent.IndexOf("## 發布版本歷史");
                    if (startIndex != -1)
                    {
                        // 找下一個主要章節的開始
                        int endIndex = readmeContent.IndexOf("## 功能特色", startIndex);
                        if (endIndex == -1)
                            endIndex = readmeContent.Length;

                        string versionHistory = readmeContent.Substring(startIndex, endIndex - startIndex).Trim();
                        versionHistory = versionHistory.Replace("###", "📌")
                                                     .Replace("✅", "✓")
                                                     .Replace("**", "")
                                                     .Replace("`", "");

                        VersionHistoryTextBlock.Text = versionHistory;
                    }
                    else
                    {
                        VersionHistoryTextBlock.Text = "未找到版本歷史資訊";
                    }
                }
                else
                {
                    VersionHistoryTextBlock.Text = "README.md 檔案不存在";
                }
            }
            catch (Exception ex)
            {
                VersionHistoryTextBlock.Text = $"讀取版本歷史時發生錯誤：{ex.Message}";
                System.Diagnostics.Debug.WriteLine($"[ERROR] LoadVersionHistory 發生錯誤: {ex}");
            }
        }
        // 統一的頁面切換方法
        private void ShowPanel(PanelType panelType)
        {
            // 隱藏所有頁面
            MainPanel.Visibility = Visibility.Collapsed;
            SettingsPanel.Visibility = Visibility.Collapsed;
            PreviewPanel.Visibility = Visibility.Collapsed;
            ConversionInfoPanel.Visibility = Visibility.Collapsed;
            VersionHistoryPanel.Visibility = Visibility.Collapsed;

            // 顯示指定頁面
            switch (panelType)
            {
                case PanelType.Main:
                    MainPanel.Visibility = Visibility.Visible;
                    break;
                case PanelType.Settings:
                    SettingsPanel.Visibility = Visibility.Visible;
                    break;
                case PanelType.Preview:
                    PreviewPanel.Visibility = Visibility.Visible;
                    break;
                case PanelType.ConversionInfo:
                    ConversionInfoPanel.Visibility = Visibility.Visible;
                    break;
                case PanelType.VersionHistory:
                    VersionHistoryPanel.Visibility = Visibility.Visible;
                    break;
            }
        }
        // "開啟檔案"相關
        private void OpenFile(string filePath)
        {
            try
            {
                // 重點：先檢查 CheckBox 是否被勾選
                if (AutoOpenFileCheckBox.IsChecked != true)
                {
                    // 沒有勾選，就不自動開啟檔案
                    return;
                }
                if (File.Exists(filePath))
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = filePath,
                        UseShellExecute = true
                    });
                    LogUsage("開啟檔案", System.IO.Path.GetFileName(filePath));
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"檔案不存在：{filePath}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"無法開啟檔案：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void AutoOpenFileCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            isAutoOpenFileEnabled = true;
            SaveAutoOpenSetting(true);
        }

        private void AutoOpenFileCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            isAutoOpenFileEnabled = false;
            SaveAutoOpenSetting(false);
        }

        private void ChooseFile_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var openFileDialog = new OpenFileDialog
                {
                    Title = "選擇檔案",
                    Filter = "Excel 檔案 (*.xlsx;*.xls;*.xlt)|*.xlsx;*.xls;*.xlt|所有檔案 (*.*)|*.*"
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    // 清空多檔案選擇相關
                    SelectedFilePaths.Clear();
                    FileDataDict.Clear();
                    UpdateSelectedFilesList();

                    FilePathTextBox.Text = openFileDialog.FileName;
                    ResetCustomerProductFields(true);
                    ReadExcelBom(openFileDialog.FileName);
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 檔案選取完成，自動載入 {BomPreviewList.Count} 筆資料");
                    AddConversionMessage($"[DEBUG] 檔案選取完成，自動載入 {BomPreviewList.Count} 筆資料");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"選擇檔案時發生錯誤：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
                System.Diagnostics.Debug.WriteLine($"[ERROR] ChooseFile_Click 發生錯誤: {ex}");
                AddConversionMessage($"[ERROR] ChooseFile_Click 發生錯誤: {ex}");
            }
        }

        private void SaveAutoOpenSetting(bool isEnabled)
        {
            try
            {
                Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\SchBomConvert",
                                                 "AutoOpenFile", isEnabled);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"儲存設定時發生錯誤：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void LoadAutoOpenSetting()
        {
            try
            {
#pragma warning disable CS8600
                object value = Microsoft.Win32.Registry.GetValue(@"HKEY_CURRENT_USER\Software\SchBomConvert",
                                                                 "AutoOpenFile", false);
#pragma warning restore CS8600
                if (value != null)
                {
                    isAutoOpenFileEnabled = Convert.ToBoolean(value);
                }
            }
            catch (Exception)
            {
                isAutoOpenFileEnabled = false;
            }
        }

        private void ReadButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string? path = FilePathTextBox.Text?.Trim();
                if (string.IsNullOrWhiteSpace(path))
                {
                    MessageBox.Show("請先選擇檔案路徑", "路徑錯誤", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (!File.Exists(path))
                {
                    MessageBox.Show("檔案不存在，請重新選擇正確的檔案", "檔案錯誤", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                ReadExcelBom(path);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"讀取檔案時發生錯誤：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
                System.Diagnostics.Debug.WriteLine($"[ERROR] ReadButton_Click 發生錯誤: {ex}");
                AddConversionMessage($"[ERROR] ReadButton_Click 發生錯誤: {ex}");
            }
        }

        private void ReadExcelBom(string fullPath)
        {
            try
            {
                string extension = Path.GetExtension(fullPath).ToLower();
                if (extension != ".xls" && extension != ".xlsx" && extension != ".xlt")
                {
                    MessageBox.Show("請選擇正確的 Excel 檔案格式 (.xls, .xlsx, .xlt)", "檔案格式錯誤", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (!File.Exists(fullPath))
                {
                    MessageBox.Show("檔案不存在，請重新選擇檔案", "檔案錯誤", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                BomList.Clear();
                BomPreviewList.Clear();
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                using var stream = File.Open(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

                IExcelDataReader reader = null;
                DataSet dataSet = null;

                try
                {
                    reader = ExcelReaderFactory.CreateReader(stream);
                    dataSet = reader.AsDataSet();

                    if (dataSet?.Tables?.Count == 0)
                    {
                        MessageBox.Show("Excel 檔案中沒有工作表", "檔案內容錯誤", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    var table = dataSet.Tables[0];

                    if (table.Rows.Count <= 5)
                    {
                        MessageBox.Show("Excel 檔案內容不足，請確認檔案格式是否正確", "檔案格式錯誤", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    var rawMainList = new List<BomPreviewItem>();
                    int processedCount = 0;
                    int errorCount = 0;

                    for (int i = 5; i < table.Rows.Count; i++)
                    {
                        try
                        {
                            var row = table.Rows[i];
                            if (row.ItemArray.All(cell => string.IsNullOrWhiteSpace(cell?.ToString())))
                                continue;

                            BomList.Add(new BomItem
                            {
                                ItemNumber = GetSafeStringValue(row, 0),
                                ShortName = GetSafeStringValue(row, 1),
                                Spec = GetSafeStringValue(row, 2),
                                Package = GetSafeStringValue(row, 3),
                                Quantity = GetSafeStringValue(row, 4),
                                PartNumber = GetSafeStringValue(row, 5),
                                UnitPrice = GetSafeStringValue(row, 6),
                                Brand = GetSafeStringValue(row, 7),
                                Note = GetSafeStringValue(row, 8),
                                Process = GetSafeStringValue(row, 9),
                                AltPart1 = GetSafeStringValue(row, 10),
                                ChinaPart = GetSafeStringValue(row, 11),
                                ChinaBrand = GetSafeStringValue(row, 12),
                                ChinaAlt = GetSafeStringValue(row, 13),
                                AltPart2 = GetSafeStringValue(row, 14)
                            });

                            string partNumber = GetSafeStringValue(row, 5);
                            if (string.IsNullOrWhiteSpace(partNumber))
                            {
                                continue;
                            }

                            var parts = partNumber.Split(new[] { ',', '，', ';', '；' }, StringSplitOptions.RemoveEmptyEntries);
                            int qty = parts.Length;

                            if (qty == 0) continue;

                            decimal price = 0;
                            string priceString = GetSafeStringValue(row, 6);
                            if (!string.IsNullOrWhiteSpace(priceString))
                            {
                                if (!decimal.TryParse(priceString, out price))
                                {
                                    System.Diagnostics.Debug.WriteLine($"[WARNING] 第 {i + 1} 行單價解析失敗: {priceString}");
                                    AddConversionMessage($"[WARNING] 第 {i + 1} 行單價解析失敗: {priceString}");
                                    price = 0;
                                }
                            }

                            // 取得規格並處理 NC 格式
                            string spec = GetSafeStringValue(row, 2);
                            // 標準化 NC 格式
                            spec = NormalizeNCFormat(spec);  

                            rawMainList.Add(new BomPreviewItem
                            {
                                IsAltLine = false,
                                PartName = GetFinalPartName(GetSafeStringValue(row, 1, "零件名稱"), partNumber),
                                Spec = spec,
                                Package = GetSafeStringValue(row, 3, "包裝"),
                                Quantity = qty,
                                Code = partNumber,
                                UnitPrice = price,
                                Vendor = GetSafeStringValue(row, 7),
                                VendorCN = string.IsNullOrWhiteSpace(GetSafeStringValue(row, 13))
                                ? GetSafeStringValue(row, 12)
                                : GetSafeStringValue(row, 13),
                                Note = GetSafeStringValue(row, 8),
                                Alt1 = GetSafeStringValue(row, 10),
                                AltCN = GetSafeStringValue(row, 13),
                                Alt2 = GetSafeStringValue(row, 14)
                            });

                            processedCount++;
                        }
                        catch (Exception rowEx)
                        {
                            errorCount++;
                            System.Diagnostics.Debug.WriteLine($"[ERROR] 處理第 {i + 1} 行時發生錯誤: {rowEx.Message}");
                            AddConversionMessage($"[ERROR] 處理第 {i + 1} 行時發生錯誤: {rowEx.Message}");

                            if (errorCount > 10)
                            {
                                MessageBox.Show($"檔案中有太多錯誤行 ({errorCount} 行)，請檢查檔案格式", "資料錯誤", MessageBoxButton.OK, MessageBoxImage.Warning);
                                break;
                            }
                        }
                    }

                    if (processedCount == 0)
                    {
                        MessageBox.Show("沒有找到有效的 BOM 資料，請檢查檔案格式和內容", "無有效資料", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    var sortedList = rawMainList
                        .OrderBy(GetMainSortKey)
                        .ToList();

                    // 合併相同零件
                    var mergedList = MergeDuplicateItems(sortedList);

                    int index = 1;
                    var grouped = mergedList
                        .OrderBy(GetCategoryGroupOrder)
                        .GroupBy(p => p.Category);

                    foreach (var group in grouped)
                    {
                        BomPreviewList.Add(new BomPreviewItem
                        {
                            IsAltLine = true,
                            PartName = group.Key.Replace(" 零件", "") + "料表",
                            Spec = "",
                            Vendor = "",
                            VendorCN = "",
                            Note = ""
                        });

                        foreach (var main in group)
                        {
                            main.Index = index++;
                            BomPreviewList.Add(main);
                            AddAltRow(main.Alt1, "替代料");
                            AddAltRow(main.Alt2, "替代料2");
                        }
                    }

                    if (errorCount > 0)
                    {
                        MessageBox.Show($"檔案載入完成！\n成功處理：{processedCount} 筆資料\n錯誤行數：{errorCount} 行",
                            "載入完成", MessageBoxButton.OK, MessageBoxImage.Information);
                        AddConversionMessage($"[INFO] 檔案載入完成！成功處理：{processedCount} 筆資料，錯誤行數：{errorCount} 行");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[INFO] 檔案載入完成，成功處理 {processedCount} 筆資料");
                        AddConversionMessage($"[INFO] 檔案載入完成，成功處理 {processedCount} 筆資料");
                    }
                    
                    void AddAltRow(string? text, string label)
                    {
                        if (string.IsNullOrWhiteSpace(text)) return;

                        try
                        {
                            //var (spec, vendor) = ParseAlternativePart(text.Trim());
                            var (spec, vendor, flagged) = AltPartRules.Parse(text.Trim());
                            BomPreviewList.Add(new BomPreviewItem
                            {
                                IsAltLine = true,
                                PartName = $"{label}:",
                                Spec = spec,
                                VendorCN = vendor,
                                Index = null
                            });
                        }
                        catch (Exception altEx)
                        {
                            System.Diagnostics.Debug.WriteLine($"[ERROR] 處理替代料時發生錯誤: {altEx.Message}");

                            BomPreviewList.Add(new BomPreviewItem
                            {
                                IsAltLine = true,
                                PartName = $"{label}:",
                                Spec = text,
                                VendorCN = "",
                                Index = null
                            });
                        }
                    }

                }
                finally
                {
                    reader?.Dispose();
                    dataSet?.Dispose();
                }
            }
            catch (UnauthorizedAccessException)
            {
                MessageBox.Show("沒有權限存取該檔案，請檢查檔案權限", "權限錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (FileNotFoundException)
            {
                MessageBox.Show("找不到指定的檔案", "檔案不存在", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (InvalidDataException)
            {
                MessageBox.Show("檔案格式不正確或已損壞，請確認這是有效的 Excel 檔案", "檔案格式錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"讀取檔案時發生未預期的錯誤：\n{ex.Message}\n\n請確認檔案格式是否正確", "讀取錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
                System.Diagnostics.Debug.WriteLine($"[ERROR] ReadExcelBom 發生錯誤: {ex}");
                AddConversionMessage($"[ERROR] ReadExcelBom 發生錯誤: {ex}");
            }
        } 

        // 標準化 NC 格式：將結尾的 " NC" 轉換為 "(NC)"
        private string NormalizeNCFormat(string spec)
        {
            if (string.IsNullOrWhiteSpace(spec))
                return spec;

            // 已經是 (NC) ，保持不變
            if (spec.Contains("(NC)", StringComparison.OrdinalIgnoreCase))
                return spec;

            // 處理結尾的 " NC"（前面有空格的 NC）
            if (spec.EndsWith(" NC", StringComparison.OrdinalIgnoreCase))
            {
                // 移除結尾的 " NC" 並加上 "(NC)"
                spec = spec.Substring(0, spec.Length - 3) + " (NC)";
            }
            // 處理開頭就是 "NC " 的情況
            else if (spec.StartsWith("NC ", StringComparison.OrdinalIgnoreCase))
            {
                spec = "(NC) " + spec.Substring(3);
            }
            // 處理中間的 " NC "
            else if (spec.Contains(" NC ", StringComparison.OrdinalIgnoreCase))
            {
                spec = Regex.Replace(spec, @"\sNC\s", " (NC) ", RegexOptions.IgnoreCase);
            }
            // 處理只有 "NC" 的情況
            else if (spec.Equals("NC", StringComparison.OrdinalIgnoreCase))
            {
                spec = "(NC)";
            }

            return spec;
        }

        // 解析替代料
        private (string spec, string vendor) ParseAlternativePart(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return ("", "");

            text = text.Trim();

            // 標準格式：規格(廠商)
            var standardMatch = Regex.Match(text, @"^(.+?)\s*\((.+?)\)\s*$");
            if (standardMatch.Success)
            {
                string potentialSpec = standardMatch.Groups[1].Value.Trim();
                string potentialVendor = standardMatch.Groups[2].Value.Trim();

                if (IsLikelyVendorName(potentialVendor) &&
                    !IsPartOfSpecification(potentialVendor) &&
                    potentialVendor.Length > 1)
                {
                    return (potentialSpec, potentialVendor);
                }
            }

            // 廠商關鍵字匹配
            var vendorKeywords = new[]
            {
                "TI", "ADI", "MAXIM", "LINEAR", "INFINEON", "VISHAY", "ROHM", "MURATA", "TDK", "IXYS", "TKS",
                "SAMSUNG", "PANASONIC", "NICHICON", "RUBYCON", "KEMET", "AVX", "YAGEO", "BOURNS", "GT(Samxon)", "Samxon",
                "DIODES", "FAIRCHILD", "ON", "NEXPERIA", "MICROCHIP", "ATMEL", "CYPRESS", "ALTERA",
                "XILINX", "LATTICE", "ANALOG", "MAXLINEAR", "BROADCOM", "MARVELL", "QUALCOMM",
                "OSRAM", "CREE", "LUMILEDS", "NICHIA", "CITIZEN", "SHARP", "TOSHIBA", "MITSUBISHI",
                "OMRON", "TYCO", "MOLEX", "JST", "HIROSE", "SAMTEC", "TE", "AMPHENOL", "FOXCONN",
                "立創", "嘉立創", "韋爾", "聖邦", "思瑞浦", "芯海", "兆易", "全志", "瑞芯微", "晶豐明源",
                "上海如韻", "矽力杰", "中穎", "華大", "敏芯", "匯頂", "卓勝微", "紫光", "海思", "展訊",
                "CORP", "CORPORATION", "TECH", "TECHNOLOGY", "SEMICONDUCTOR", "SEMI", "ELECTRONICS", "Comchip",
                "littlefuse", "JTC" , "PFC" ,
                "凱恩傑" , "西五" , "營格" ,
                "耕興", "友士", "航興", "朝欣", "松川", "九寧", "奇普仕", "新承", "緯澄", "安帝", "晟通", "功得", "偉強",
                "代理", "原廠", "官方", "授權"
            };

            var foundVendor = vendorKeywords.FirstOrDefault(keyword =>
                text.Contains(keyword, StringComparison.OrdinalIgnoreCase));

            if (foundVendor != null)
            {
                return ExtractVendorAndSpec(text, foundVendor);
            }

            return TrySeparatorSplit(text);
        }

        // 額外提取廠商和規格
        private (string spec, string vendor) ExtractVendorAndSpec(string text, string foundVendor)
        {
            int vendorIndex = text.LastIndexOf(foundVendor, StringComparison.OrdinalIgnoreCase);
            string vendorPart = text.Substring(vendorIndex).Trim();
            string spec = text.Substring(0, vendorIndex).Trim();

            string vendor;
            if (foundVendor.Length >= 2 && IsChineseName(foundVendor))
            {
                vendor = foundVendor;
            }
            else
            {
                var words = vendorPart.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                vendor = words.Length > 1 && IsLikelyVendorExtension(words[1].ToUpperInvariant())
                    ? $"{words[0]} {words[1]}"
                    : words[0];
            }

            return (CleanSpecString(spec), vendor);
        }

        // 分隔符分割
        private (string spec, string vendor) TrySeparatorSplit(string text)
        {
            string[] separators = { "/", " ", ":", "：", "，", ";", "；", "|", "\\", "　" };

            foreach (string separator in separators)
            {
                if (!text.Contains(separator)) continue;

                var parts = text.Split(new[] { separator }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length < 2) continue;

                string firstPart = parts[0].Trim();
                string lastPart = parts[parts.Length - 1].Trim();

                if (IsLikelyVendorName(lastPart) && IsLikelyPartNumber(firstPart))
                {
                    return (string.Join(" ", parts.Take(parts.Length - 1).Select(p => p.Trim())), lastPart);
                }
                if (IsLikelyVendorName(firstPart) && IsLikelyPartNumber(parts[1].Trim()))
                {
                    return (string.Join(" ", parts.Skip(1).Select(p => p.Trim())), firstPart);
                }
            }

            // 預設處理
            return IsLikelyPartNumber(text) ? (text, "") :
                   IsLikelyVendorName(text) ? ("", text) :
                   (text, "");
        }

        // 檢查是否為規格的一部分
        private bool IsPartOfSpecification(string text) =>
            new[] { "Samxon", "Samx", "CapXon", "Rubycon", "Panasonic", "Nichicon", "MST",
            "United", "Lelon", "GT", "LZ", "KZE", "KZH", "UPW", "UPS",
            "ESR", "Low", "High", "Temp", "V", "uF", "nF", "pF" }
            .Any(marker => text.Contains(marker, StringComparison.OrdinalIgnoreCase)) || text.Length <= 6;

        // 檢查是否為中文名稱
        private bool IsChineseName(string text) =>
            Regex.IsMatch(text, @"[\u4e00-\u9fa5]");

        // 檢查是否為廠商延伸名稱
        private bool IsLikelyVendorExtension(string text) =>
            new[] { "TECH", "TECHNOLOGY", "SEMICONDUCTOR", "ELECTRONICS", "CORP", "CORPORATION", "LTD", "LIMITED", "INC" }
            .Any(ext => text.Equals(ext, StringComparison.OrdinalIgnoreCase));

        // 清理規格字串
        private string CleanSpecString(string spec)
        {
            string cleaned = spec.Trim(' ', '-', '/', ':', '：', '，', ';', '；', '|', '\\', '　');

            // 統一所有空白為單空格
            cleaned = Regex.Replace(cleaned, @"\s+", " ");

            // 在SMD尺寸後面加雙空格
            cleaned = Regex.Replace(cleaned, @"\b(0402|0603|0805|1206|1210|2010|1812|2512|2728|3920|SR3920)\b\s*", "$1  ");

            // 數值單位組合後面加雙空格
            cleaned = Regex.Replace(cleaned, @"\b(\d+\.?\d*[RKMΩpnuμmkMGTHz]{1,3}F?)\b\s*", "$1  ");

            // 在百分比後面加雙空格
            cleaned = Regex.Replace(cleaned, @"\b(\d+\.?\d*%)\b\s*", "$1  ");

            // 清理結尾多餘空格
            cleaned = cleaned.TrimEnd();

            return cleaned;
        }

        // 檢查是否為零件編號
        private bool IsLikelyPartNumber(string text) =>
            !string.IsNullOrWhiteSpace(text) &&
            Regex.IsMatch(text, @"[A-Za-z]\d+|[A-Za-z]+\d+[A-Za-z]*|\d+[A-Za-z]+|[A-Za-z]+\d+[-_\.]*[A-Za-z\d]*") &&
            !Regex.IsMatch(text, @"^[\u4e00-\u9fa5]+$") &&
            text.Length <= 30;

        // 檢查是否為廠商名稱
        private bool IsLikelyVendorName(string text)
        {
            if (string.IsNullOrWhiteSpace(text) || IsPartOfSpecification(text)) return false;

            var companyKeywords = new[]
            {
                "科技", "電子", "半導體", "有限", "公司", "股份", "集團", "代理", "原廠",
                "TECH", "TECHNOLOGY", "SEMICONDUCTOR", "ELECTRONICS", "CORP", "CORPORATION",
                "LIMITED", "LTD", "CO", "INC", "GROUP"
            };

            return companyKeywords.Any(keyword => text.Contains(keyword, StringComparison.OrdinalIgnoreCase)) ||
                   Regex.IsMatch(text, @"^[A-Z]{2,10}$") ||
                   (Regex.IsMatch(text, @"[\u4e00-\u9fa5]") && text.Length <= 20 && text.Length >= 2);
        }

        private string GetBestPackageValue(IEnumerable<string?> packages)
        {
            var validPackages = packages
                .Where(pkg => !string.IsNullOrWhiteSpace(pkg))
                .Select(pkg => pkg!.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (validPackages.Count == 0)
                return "";

            if (validPackages.Count == 1)
                return validPackages.First();

            // 如果有多個不同的包裝值，優先選擇包含尺寸資訊的
            var smdSizes = new[] { "0402" , "0603", "0805", "1206", "1210", "2010" , "1812", "2512", "2728", "3920" , "SR3920" };
            foreach (var size in smdSizes)
            {
                var matchingPackage = validPackages.FirstOrDefault(pkg =>
                    pkg.Contains(size, StringComparison.OrdinalIgnoreCase));
                if (matchingPackage != null)
                    return matchingPackage;
            }
            // 其他封裝類型
            var packageTypes = new[] { "DIP", "SOP", "QFP", "BGA", "PLCC", "SOIC", "TSSOP", "SOT", "TO" };
            foreach (var type in packageTypes)
            {
                var matchingPackage = validPackages.FirstOrDefault(pkg =>
                    pkg.Contains(type, StringComparison.OrdinalIgnoreCase));
                if (matchingPackage != null)
                    return matchingPackage;
            }
            return validPackages.First();
        }
        private List<BomPreviewItem> MergeDuplicateItems(List<BomPreviewItem> items)
        {
            var mergedItems = new List<BomPreviewItem>();

            // 分組時包含 Package 欄位（保留尺寸資訊）
            var groupedItems = items
                .GroupBy(item => new
                {
                    PartName = item.PartName?.Trim(),
                    NormalizedSpec = NormalizeSpec(item.Spec),
                    Package = item.Package?.Trim(),  // 加入 Package 分類方法區分不同尺寸
                    UnitPrice = item.UnitPrice
                })
                .ToList();

            foreach (var group in groupedItems)
            {
                if (group.Count() == 1)
                {
                    mergedItems.Add(group.First());
                }
                else
                {
                    var firstItem = group.First();
                    var mergedItem = new BomPreviewItem
                    {
                        IsAltLine = firstItem.IsAltLine,
                        PartName = firstItem.PartName,
                        Spec = GetBestSpecValue(group.Select(g => g.Spec)),
                        Package = firstItem.Package,  // 保留原始 Package
                        UnitPrice = firstItem.UnitPrice,
                        Vendor = firstItem.Vendor,
                        VendorCN = firstItem.VendorCN,
                        Note = MergeNotes(group.Select(g => g.Note)),
                        Quantity = group.Sum(g => g.Quantity),
                        Code = MergePartNumbers(group.Select(g => g.Code)),
                        Alt1 = MergeAlternatives(group.Select(g => g.Alt1)),
                        Alt2 = MergeAlternatives(group.Select(g => g.Alt2)),
                        AltCN = MergeAlternatives(group.Select(g => g.AltCN))
                    };

                    mergedItems.Add(mergedItem);
                    string msg = $"[INFO] 合併零件: {firstItem.PartName} ({firstItem.Spec}) 尺寸:{firstItem.Package} - 合併了 {group.Count()} 項，總數量: {mergedItem.Quantity}";
                    System.Diagnostics.Debug.WriteLine(msg);
                    AddConversionMessage(msg);
                }
            }

            return mergedItems;
        }
        // 標準化規格（移除尺寸等差異）
        private string NormalizeSpec(string? spec)
        {
            if (string.IsNullOrWhiteSpace(spec)) return "";

            string normalized = spec.Trim();

            // 移除常見的尺寸資訊，保留核心規格
            var smdSizes = new[] { "0402" , "0603", "0805", "1206", "1210" , "2010" , "1812" , "2512", "2728", "3920" , "SR3920" };
            foreach (var size in smdSizes)
            {
                normalized = normalized.Replace(size, "", StringComparison.OrdinalIgnoreCase);
            }

            // 清理多餘空格
            normalized = System.Text.RegularExpressions.Regex.Replace(normalized, @"\s+", " ").Trim();

            return normalized.ToUpper();
        }

        // 選擇最完整的規格值
        private string GetBestSpecValue(IEnumerable<string?> specs)
        {
            var validSpecs = specs
                .Where(spec => !string.IsNullOrWhiteSpace(spec))
                .Select(spec => spec!.Trim())
                .OrderByDescending(spec => spec.Length) // 選擇最長的（通常最完整）
                .ToList();

            return validSpecs.FirstOrDefault() ?? "";
        }

        // 合併零件編號
        private string MergePartNumbers(IEnumerable<string?> partNumbers)
        {
            var validNumbers = partNumbers
                .Where(pn => !string.IsNullOrWhiteSpace(pn))
                .SelectMany(pn => pn!.Split(new[] { ',', '，', ';', '；' }, StringSplitOptions.RemoveEmptyEntries))
                .Select(pn => pn.Trim())
                .Where(pn => !string.IsNullOrWhiteSpace(pn))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(pn => GetPartNumberSortKey(pn))
                .ThenBy(pn => pn) // 同號碼時照字母順序
                .ToList();

            return string.Join(", ", validNumbers);
        }

        private int GetPartNumberSortKey(string pn)
        {
            // 只抓第一個數字
            var match = System.Text.RegularExpressions.Regex.Match(pn, @"[A-Za-z]+(\d+)");
            if (match.Success && int.TryParse(match.Groups[1].Value, out int num))
                return num;
            return int.MaxValue;
        }

        // 合併備註
        private string MergeNotes(IEnumerable<string?> notes)
        {
            var validNotes = notes
                .Where(note => !string.IsNullOrWhiteSpace(note))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            return validNotes.Count > 0 ? string.Join("; ", validNotes) : "";
        }

        // 合併替代料
        private string MergeAlternatives(IEnumerable<string?> alternatives)
        {
            var validAlts = alternatives
                .Where(alt => !string.IsNullOrWhiteSpace(alt))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            return validAlts.Count > 0 ? string.Join("; ", validAlts) : "";
        }
        // 決定正式零件名稱
        private string GetFinalPartName(string originalPartName, string partNumber)
        {
            // 如果原始零件名稱存在，直接使用
            if (!string.IsNullOrWhiteSpace(originalPartName))
                return originalPartName;

            // 嘗試自動識別
            string autoDetected = AutoDetectPartNameFromCode(partNumber);

            // 如果自動識別成功，使用識別結果
            if (!string.IsNullOrWhiteSpace(autoDetected))
                return autoDetected;

            // 如果都沒有，保持空白，讓系統後續根據其他資訊判斷
            return "";
        }

        // 根據零件編號自動識別零件類型 僅出現在沒有打上"零件名稱"的時候
        private string AutoDetectPartNameFromCode(string partNumber)
        {
            if (string.IsNullOrWhiteSpace(partNumber)) return "SMD 元件";

            // 使用不同的變數名稱避免衝突
            var partCodes = partNumber.Split(new[] { ',', '，', ';', '；' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var partCode in partCodes)
            {
                string trimmedCode = partCode.Trim().ToUpper();

                // 檢查是否以特定前綴開始
                if (trimmedCode.StartsWith("ZD"))
                {
                    return "SMD 稽納";
                }
                if (trimmedCode.StartsWith("BAT"))
                {
                    return "電池";
                }

                
                //這幾項檢測太籠統 等找到方法在放回去

                /*
                else if (trimmedCode.StartsWith("D") && !trimmedCode.StartsWith("DC"))
                {
                    return "SMD 二極體";
                }
                else if (trimmedCode.StartsWith("R"))  
                {
                    return "SMD 電阻";
                }
                else if (trimmedCode.StartsWith("L"))
                {
                    return "SMD 電感";
                }
                else if (trimmedCode.StartsWith("Q"))
                {
                    return "SMD 電晶體";
                }
                else if (trimmedCode.StartsWith("U") || trimmedCode.StartsWith("IC"))
                {
                    return "SMD IC";
                }
                */

                // 可以擴展其他規則...
            }

            return "";
        }
        private string GetSafeStringValue(DataRow row, int columnIndex, string fieldType = "")
        {
            try
            {
                string value = row?[columnIndex]?.ToString()?.Trim() ?? string.Empty;

                if (string.IsNullOrEmpty(value))
                    return string.Empty;

                // 清理 [NoValue] 
                value = CleanNoValueString(value);

                if (fieldType.ToLower().Contains("package") || fieldType.Contains("包裝"))
                {
                    return value;
                }

                // 只對特定欄位進行統一處理
                if (ShouldUnifyField(fieldType))
                {
                    var parts = value.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                                   .Select(p => p.Trim())
                                   .Where(p => !string.IsNullOrEmpty(p))
                                   .ToArray();

                    return UnifyMultipleValues(parts, fieldType);
                }

                return value;
            }
            catch { return string.Empty; }
        }

        // 加入欄位類型參數
        private string UnifyMultipleValues(string[] values, string fieldType = "")
        {
            if (values == null || values.Length == 0)
                return string.Empty;

            // 如果是零件名稱欄位，強制移除所有尺寸資訊
            if (fieldType.ToLower().Contains("partname") || fieldType.Contains("零件名稱"))
            {
                return GetCleanComponentName(values);
            }

            // 如果是包裝欄位，優先返回尺寸
            if (fieldType.ToLower().Contains("package") || fieldType.Contains("包裝"))
            {
                return GetPackageInfo(values);
            }

            // 預設處理（向後相容）
            return GetCleanComponentName(values);
        }

        private string GetCleanComponentName(string[] values)
        {
            // 元件分類映射改由 ComponentMappingService 載入 (component_mappings.json)
            // 順序敏感：具體類型擺前面 (例：合金電阻 → 電阻)；UI 可在「🧩 元件映射」管理
            var componentMappings = ComponentMappingService.Instance.Mappings;

            // 組合所有值進行分析
            string combinedText = string.Join(" ", values).ToLower();

            // 按優先級檢查元件類型 - 先檢查更具體的類型
            foreach (var mapping in componentMappings)
            {
                string targetName = mapping.Name;
                var keywords = mapping.Keywords;

                bool hasKeyword = keywords.Any(keyword =>
                    combinedText.Contains(keyword.ToLower()));

                if (hasKeyword)
                {
                    // 特別處理DIP元件 - 需要同時包含DIP關鍵字
                    if (targetName.StartsWith("DIP"))
                    {
                        bool isDip = combinedText.Contains("dip");
                        if (isDip) return targetName;
                    }
                    // 特別處理SMD元件 - 優先識別SMD
                    else if (targetName.StartsWith("SMD"))
                    {
                        bool isSmd = combinedText.Contains("smd") ||
                                   HasSmdSizeIndicator(combinedText);
                        if (isSmd) return targetName;

                        // 如果有SMD尺寸但沒有明確SMD關鍵字，也視為SMD
                        if (HasSmdSizeIndicator(combinedText)) return targetName;
                    }
                    else
                    {
                        return targetName;
                    }
                }
            }

            // 沒有找到任何匹配時，才使用預設值
            // 如果包含SMD但沒有匹配到具體類型
            if (combinedText.Contains("smd") || HasSmdSizeIndicator(combinedText))
            {
                return "SMD 元件";
            }

            // 如果包含DIP但沒有匹配到具體類型  
            if (combinedText.Contains("dip"))
            {
                return "DIP 元件";
            }

            // 移除所有尺寸後返回第一個值
            return RemoveAllSizes(values[0]);
        }

        // 檢查是否包含SMD尺寸指示器
        private bool HasSmdSizeIndicator(string text)
        {
            string[] smdSizes = { "0402", "0603", "0805", "1206", "1210", "2010", "1812", "2512", "2728", "3920", "SR3920" };
            return smdSizes.Any(size => text.Contains(size));
        }

        // 獲取封裝資訊
        private string GetPackageInfo(string[] values)
        {
            string[] smdSizes = { "0402", "0603", "0805", "1206", "1210", "2010", "1812", "2512", "2728", "3920", "SR3920" };
            foreach (string size in smdSizes)
            {
                var foundSize = values.FirstOrDefault(v => v.Contains(size));
                if (foundSize != null) return size;
            }

            // 檢查其他封裝類型
            string[] packageTypes = { "DIP", "SOP", "QFP", "BGA", "PLCC", "SOIC", "TSSOP", "SOT", "TO" };
            foreach (string packageType in packageTypes)
            {
                var foundPackage = values.FirstOrDefault(v =>
                    v.Contains(packageType, StringComparison.OrdinalIgnoreCase));
                if (foundPackage != null) return packageType;
            }

            return RemoveAllSizes(values[0]);
        }

        // 強力移除所有尺寸資訊
        private string RemoveAllSizes(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;

            string result = input;

            // 移除所有可能的尺寸格式
            string[] patterns = {
                @"\b\d{4}\b",           // 4位數字 (0805, 1206等)
                @"\b\d{3}\b",           // 3位數字
                @"\d+\s*x\s*\d+",       // NxN格式
                @"\d+\.\d+\s*mm",       // N.N mm格式
                @"\d+mm",               // Nmm格式
                @"SOT-\d+",             // SOT-XX格式
                @"TO-\d+",              // TO-XX格式
            };

            foreach (string pattern in patterns)
            {
                result = Regex.Replace(result, pattern, "", RegexOptions.IgnoreCase);
            }

            // 清理多餘的空格和標點
            result = Regex.Replace(result, @"\s+", " ");
            result = result.Trim(' ', '-', '_', ',', ';', '/', '\\');

            return result;
        }
        private string GetSafeStringValue(DataRow row, int columnIndex)
{
    return GetSafeStringValue(row, columnIndex, "");
}

// 判斷是否需要統一的欄位
private bool ShouldUnifyField(string fieldType)
{
    return fieldType.ToLower() switch
    {
        "partname" or "零件名稱" or "package" or "包裝" => true,
        _ => false
    };
}

// 單純清理所有無效值的方法（加上去重複）
private string CleanNoValueString(string input)
{
    if (string.IsNullOrWhiteSpace(input))
        return string.Empty;
    
    // 定義所有需要清理的無效值模式
    string[] invalidPatterns = 
    { 
        "[NoValue]", "[NO VALUE]", "[NOVALUE]",
        "[NoParam]", "[NO PARAM]", "[NOPARAM]",
        "[NULL]", "[EMPTY]", "[NONE]", "[N/A]",
        "NoValue", "NO VALUE", "NOVALUE",
        "NoParam", "NO PARAM", "NOPARAM",
        "NULL", "EMPTY", "NONE", "N/A"
    };
    
    var parts = input.Split(new[] { ',', ';', '|' }, StringSplitOptions.RemoveEmptyEntries)
                     .Select(p => p.Trim())
                     .Where(p => !string.IsNullOrEmpty(p) && 
                                !IsInvalidValue(p, invalidPatterns))
                     .Distinct(StringComparer.OrdinalIgnoreCase) // 去除重複值
                     .ToArray();
    
    return parts.Length > 0 ? string.Join(", ", parts) : string.Empty;
}

// 檢查是否為無效值
private bool IsInvalidValue(string value, string[] invalidPatterns)
{
    return invalidPatterns.Any(pattern => 
        value.Equals(pattern, StringComparison.OrdinalIgnoreCase) ||
        value.Contains(pattern, StringComparison.OrdinalIgnoreCase));
}

        // 主排序鍵：單檔匯出與合併匯出共用，避免兩處規則「改一邊漏一邊」
        // （本體由原 OrderBy lambda 逐字搬移，行為不變）
        private static string GetMainSortKey(BomPreviewItem p)
        {
            if (p.Category == "SMD 零件")
            {
                var group = GetSortGroup(p);
                return $"A-{group.sizeOrder:D3}-{group.typeOrder:D2}-{GetSortWeight(p)}";
            }
            else if (p.Category == "DIP 零件")
            {
                if ((p.PartName?.Contains("電解電容") ?? false) || (p.PartName?.Contains("金屬皮膜電容") ?? false))
                    return $"B-04-{GetCapacitorSortKey(p.Spec)}-{p.PartName}";
                if (p.PartName?.Contains("端子座") ?? false)
                    return $"B-34-{GetTerminalSortKey(p)}-{p.PartName}";   // 端子座/連接器：族群分塊→組內依腳位
                return $"B-{GetDipPartOrder(p):D2}-{p.PartName}";
            }
            else
            {
                return $"Z-{p.PartName}";
            }
        }

        // 大分類顯示順序：SMD → DIP → 其他（單檔/合併匯出共用）
        private static int GetCategoryGroupOrder(BomPreviewItem p)
            => p.Category == "SMD 零件" ? 0 : p.Category == "DIP 零件" ? 1 : 2;

        private static int GetDipPartOrder(BomPreviewItem p)
        {
            string name = p.PartName ?? "";
            string spec = p.Spec ?? "";

            int dipOrder =
                // DIP 基本元件
                name.Contains("電阻") && !name.Contains("合金電阻") ? 1 :
                name.Contains("水泥電阻") ? 2 :
                name.Contains("Y電容") ? 3 :
                name.Contains("X電容") ? 3 :
                name.Contains("電解電容") ? 4 :
                name.Contains("安規電容") ? 5 :
                name.Contains("電容") ? 6 :
                name.Contains("電感") ? 7 :
                name.Contains("CHOKE") ? 8 :
                name.Contains("變壓器") ? 9 :

                // DIP 二極體類
                name.Contains("橋式") ? 10 :
                name.Contains("二極體") ? 11 :
                name.Contains("蕭特基") ? 11 :
                name.Contains("稽納") ? 11 :
                name.Contains("LED") ? 12 :

                // DIP 保護元件
                name.Contains("保險絲") ? 13 :
                name.Contains("突波") ? 14 :
                name.Contains("NTC") ? 15 :
                name.Contains("雷擊保護") ? 16 :

                // DIP 主動元件
                name.Contains("BJT") ? 17 :
                name.Contains("MOS") ? 18 :
                name.Contains("MOSFET") ? 18 :
                name.Contains("IGBT") ? 19 :
                name.Contains("電晶體") ? 20 :
                name.Contains("RELAY") ? 21 :
                name.Contains("繼電器") ? 21 :

                // DIP IC 類
                name.Contains("MCU") ? 22 :
                name.Contains("OP") ? 23 :
                name.Contains("IPM") ? 24 :
                name.Contains("IC") ? 25 :
                name.Contains("晶片") ? 25 :

                // DIP 模組/感測器
                name.Contains("模組") ? 26 :
                name.Contains("模塊") ? 26 :
                name.Contains("感測器") ? 27 :
                name.Contains("SENSOR") ? 27 :

                // DIP 人機介面
                name.Contains("開關") ? 28 :
                name.Contains("按鍵") ? 29 :
                name.Contains("BUZZER") ? 30 :
                name.Contains("蜂鳴器") ? 30 :
                name.Contains("DISPLAY") ? 31 :
                name.Contains("顯示器") ? 31 :

                // DIP 連接器
                name.Contains("排針") ? 32 :
                name.Contains("排母") ? 33 :
                name.Contains("端子座") ? 34 :                       // 細排移至 GetMainSortKey 端子座特例（GetTerminalSortKey）
                name.Contains("端子") ? 39 :
                name.Contains("歐規端子座") ? 39 :
                name.Contains("接線座") ? 39 :

                // DIP 機構件
                name.Contains("散熱片") ? 40 :
                name.Contains("HEATSINK") ? 40 :
                name.Contains("銅塊") ? 41 :
                name.Contains("銅柱") ? 42 :
                name.Contains("銅排") ? 43 :
                name.Contains("五金") ? 44 :
                name.Contains("螺絲") ? 45 :

                // 特殊處理：根據規格判斷
                spec.Contains("uF") ? 6 :      // 電容類
                spec.Contains("MOV") ? 14 :    // 突波吸收器
                99;  // 未分類

            return dipOrder;
        }

        // 端子座/連接器排序子鍵：族群分塊 → 組內依腳位（取代舊 GetTerminalBlockOrder 系列）
        private static string GetTerminalSortKey(BomPreviewItem p)
        {
            // 型號可能在「規格」或「零件編號」欄；依序嘗試，取第一個能抓到腳位的欄位
            foreach (var field in new[] { p.Spec, p.Code, p.PartName })
            {
                if (string.IsNullOrWhiteSpace(field)) continue;
                var (family, pins) = ParseTerminal(field);
                if (pins != 999) return $"{family}-{pins:D3}";
            }
            // 都沒抓到腳位：用規格的 family、腳位排該族群最後
            var (fam, _) = ParseTerminal(p.Spec ?? p.Code ?? p.PartName ?? "");
            return $"{fam}-999";
        }

        // 通用解析：回傳 (族群, 腳位數)。腳位依序嘗試多種樣式；抓不到 → 該族群內排最後。
        private static (string family, int pins) ParseTerminal(string raw)
        {
            const int MAX = 999;
            if (string.IsNullOrWhiteSpace(raw)) return ("ZZZZ", MAX);
            string t = raw.ToUpperInvariant().Trim();

            // 規則1：開頭單一字母+數字（P4、M7）→ family=字母、pins=數字
            var lead = Regex.Match(t, @"^([A-Z])(\d{1,3})\b");
            if (lead.Success)
                return (lead.Groups[1].Value, int.Parse(lead.Groups[2].Value));

            // family：開頭連續英數（停在 '-' 或空白），例：ME040 / MX122 / CDS19 / 2530H
            var fam = Regex.Match(t, @"^([A-Z0-9]+)");
            string family = fam.Success ? fam.Groups[1].Value : "ZZZZ";

            // 規則2：歐規 -<節距>-<腳位>[P/G]，例：-508-4P、-500-2G
            var euro = Regex.Match(t, @"-\d{3,4}-(\d{1,3})\s*[PG]\b");
            if (euro.Success) return (family, int.Parse(euro.Groups[1].Value));

            // 規則3：字尾 <數字>(P|PIN|G)，例：07P、2G、4PIN
            var tail = Regex.Match(t, @"(\d{1,3})\s*(?:PIN|P|G)\b");
            if (tail.Success) return (family, int.Parse(tail.Groups[1].Value));

            return (family, MAX);
        }

        // 電阻排序方法 - 按照數字大小排序
        private static string GetResistorSortKey(string spec)
        {
            if (string.IsNullOrWhiteSpace(spec)) return "999999Z";

            // 提取電阻值和單位
            var valueMatch = Regex.Match(spec, @"([\d\.]+)\s*(R|K|M|Ω|ohm)", RegexOptions.IgnoreCase);
            if (!valueMatch.Success) return "999999Z";

            double numericValue = double.Parse(valueMatch.Groups[1].Value);
            string unit = valueMatch.Groups[2].Value.ToUpper();

            // 檢查精度
            bool hasPrecision = spec.Contains("%");
            string precisionSuffix = hasPrecision ? "1" : "0";

            // 單位排序權重：R=1, K=2, M=3
            int unitWeight = unit switch
            {
                "R" or "Ω" or "OHM" => 1,
                "K" => 2,
                "M" => 3,
                _ => 1
            };

            // 格式化：單位權重 + 數值(12位，含小數) + 精度標記
            // 單位權重排在最前 → 依實際阻值排序（R 段 → K 段 → M 段，各段內再按數值；M 才會排到最後）
            return $"{unitWeight}-{numericValue:000000000000.000}-{precisionSuffix}";
        }
        // 電容排序方法
        private static string GetCapacitorSortKey(string spec)
        {
            if (string.IsNullOrWhiteSpace(spec)) return "ZZZZZZ";

            // 先嘗試解析3位數字編碼（如104, 101等）
            var codeMatch = Regex.Match(spec, @"\b(\d{3})\b");
            if (codeMatch.Success)
            {
                string code = codeMatch.Groups[1].Value;
                double capacitance = ConvertCodeToCapacitance(code);

                // 提取電壓和材質資訊用於次要排序
                string voltage = ExtractVoltage(spec);
                string material = ExtractMaterial(spec);

                return $"{capacitance:000000000000.000}-{voltage}-{material}";
            }

            // 直接數值 + 單位的格式（如10p, 22p, 1n, 1u等）
            var directMatch = Regex.Match(spec, @"([\d\.]+)\s*(p|pF|n|nF|u|uF|μF)", RegexOptions.IgnoreCase);
            if (directMatch.Success)
            {
                double value = double.Parse(directMatch.Groups[1].Value);
                string unit = directMatch.Groups[2].Value.ToLower();

                // 轉換為pF（皮法）
                double capacitanceInPF = unit switch
                {
                    "p" or "pf" => value,
                    "n" or "nf" => value * 1_000,
                    "u" or "uf" or "μf" => value * 1_000_000,
                    _ => value
                };

                string voltage = ExtractVoltage(spec);
                string material = ExtractMaterial(spec);

                return $"{capacitanceInPF:000000000000.000}-{voltage}-{material}";
            }

            return $"ZZZZZZ-{spec}";
        }

        // 將3位數字編碼轉換為電容值（pF）
        private static double ConvertCodeToCapacitance(string code)
        {
            if (code.Length != 3) return 0;

            // 前兩位是有效數字，第三位是10的次方
            double significantDigits = double.Parse(code.Substring(0, 2));
            int multiplier = int.Parse(code.Substring(2, 1));

            return significantDigits * Math.Pow(10, multiplier);
        }

        // 提取電壓資訊
        private static string ExtractVoltage(string spec)
        {
            var voltageMatch = Regex.Match(spec, @"(\d+)V", RegexOptions.IgnoreCase);
            if (voltageMatch.Success)
            {
                return voltageMatch.Groups[1].Value.PadLeft(3, '0'); // 補零對齊
            }
            return "000"; // 沒有電壓資訊時使用000
        }

        // 提取材質資訊
        private static string ExtractMaterial(string spec)
        {
            var materialMatch = Regex.Match(spec, @"(X\d+R|X\d+S|Y\d+R|Y\d+S|NP0|C0G)", RegexOptions.IgnoreCase);
            if (materialMatch.Success)
            {
                return materialMatch.Groups[1].Value.ToUpper();
            }
            return "ZZZ"; // 沒有材質資訊時使用ZZZ
        }

        private static (int sizeOrder, int typeOrder) GetSortGroup(BomPreviewItem p)
        {
            string part = p.PartName ?? "";
            string size = ExtractSMDSize(p.Spec ?? "");
            int sizeIndex = GetSizeOrder(size);

            int typeOrder =
                // 保持電阻和電容的原有排序（它們有特殊的 GetSortWeight 處理）
                part.Contains("電阻") && !part.Contains("合金電阻") ? 0 :
                part.Contains("電容") ? 1 :

                part.Contains("合金電阻") ? 2 :
                part.Contains("鉭電") ? 3 :
                part.Contains("鋁電") ? 3 :   // 鋁電解電容：歸電解電容群，weight 走容值→電壓（見 GetSortWeight）

                part.Contains("電感") ? 4 :
                part.Contains("三端濾波器") ? 5 :

                part.Contains("二極體") ? 6 :
                part.Contains("蕭特基") ? 6 :
                part.Contains("稽納") ? 6 :
                part.Contains("LED") ? 7 :

                part.Contains("電晶體") ? 8 :
                part.Contains("MOSFET") ? 8 :

                part.Contains("保險絲") ? 9 :
                part.Contains("BEAD") ? 9 :
                part.Contains("TVS") ? 9 :
                part.Contains("ESD") ? 9 :
                part.Contains("保護元件") ? 10 :

                part.Contains("MCU") ? 11 :
                part.Contains("OP") ? 12 :
                part.Contains("DC/DC") ? 12 :
                part.Contains("LDO") ? 12 :
                part.Contains("光耦合") ? 12 :
                part.Contains("IC") ? 13 :

                part.Contains("NTC") ? 14 :
                part.Contains("IR") ? 14 :
                part.Contains("霍爾") ? 14 :
                part.Contains("CT") ? 14 :
                part.Contains("SENSOR") ? 15 :

                part.Contains("按鍵") ? 16 :
                part.Contains("開關") ? 16 :
                part.Contains("端子") ? 16 :
                part.Contains("端子座") ? 16 :
                part.Contains("晶振") ? 17 :
                part.Contains("銅塊") ? 18 :
                part.Contains("跳線") ? 19 :
                part.Contains("銅柱") ? 20 :

                99;

            return (sizeIndex, typeOrder);
        }
        // GetSortWeight方法中的電阻和電容部分
        private static string GetSortWeight(BomPreviewItem p)
        {
            string name = p.PartName ?? "";
            string spec = p.Spec ?? "";

            if (name.Contains("電阻")) return GetResistorSortKey(spec);
            if (name.Contains("電容")) return GetCapacitorSortKey(spec);
            if (name.Contains("鋁電")) return $"AL-{GetCapacitorSortKey(spec)}"; // 鋁電解：依容值→電壓
            if (name.Contains("鉭電")) return $"TA-{GetCapacitorSortKey(spec)}"; // 鉭質電容：依容值→電壓（與鋁電同 typeOrder，前綴分塊避免交錯）

            // 其他元件的排序邏輯保持不變
            if (name.Contains("電感")) return $"L-{GetSizeOrder(ExtractSMDSize(spec)):D3}";
            if (name.Contains("二極體")) return $"D-{GetSizeOrder(ExtractSMDSize(spec)):D3}";
            if (name.Contains("電晶體")) return $"Q-{GetSizeOrder(ExtractSMDSize(spec)):D3}";
            if (name.Contains("IC")) return $"U-{GetSizeOrder(ExtractSMDSize(spec)):D3}";
            if (name.Contains("感測")) return $"S-{GetSizeOrder(ExtractSMDSize(spec)):D3}";
            if (name.Contains("保護")) return $"F-{GetSizeOrder(ExtractSMDSize(spec)):D3}";

            return $"Z-{GetSizeOrder(ExtractSMDSize(spec)):D3}";
        }

        private static int GetSizeOrder(string size)
        {
            string[] sizeOrder = { "0402", "0603", "0805", "1206", "1210", "2010", "1812", "2512", "2728", "3920", "SR3920" };
            int index = Array.IndexOf(sizeOrder, size);
            return index >= 0 ? index : 999;
        }

        private static string ExtractSMDSize(string spec)
        {
            var match = Regex.Match(spec, @"\b(0402|0603|0805|1206|1210|2010|1812|2512|2728|3920|SR3920)\b");
            return match.Success ? match.Value : "ZZZZ";
        }

        private void PreviewButton_Click(object sender, RoutedEventArgs e)
        {
            ShowPanel(PanelType.Preview);
            try
            {
                string? path = FilePathTextBox.Text?.Trim();
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                {
                    MessageBox.Show("請先選擇正確的檔案路徑", "檔案錯誤", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                ReadExcelBom(path);

                if (BomPreviewList.Count == 0)
                {
                    MessageBox.Show("預覽資料為空，可能是 Excel 沒內容、格式錯誤或都被過濾了。", "預覽失敗", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                PreviewGrid.ItemsSource = BomPreviewList;
                MainPanel.Visibility = Visibility.Collapsed;
                PreviewPanel.Visibility = Visibility.Visible;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"預覽檔案時發生錯誤：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
                System.Diagnostics.Debug.WriteLine($"[ERROR] PreviewButton_Click 發生錯誤: {ex}");
                AddConversionMessage($"[ERROR] PreviewButton_Click 發生錯誤: {ex}");
            }
        }

        private void BackToMain_Click(object sender, RoutedEventArgs e)
        {
            ShowPanel(PanelType.Main);

        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            ShowPanel(PanelType.Settings);

        }

        private void ComponentMappingButton_Click(object sender, RoutedEventArgs e)
        {
            var win = new ComponentMappingWindow { Owner = this };
            win.ShowDialog();
        }
        private void SaveCustomerProductSettings()
        {
            try
            {
                Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\SchBomConvert",
                                                 "CustomerName", CustomerName ?? "");
                Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\SchBomConvert",
                                                 "ProductName", ProductName ?? "");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"儲存客戶產品設定失敗: {ex.Message}");
            }
        }

        private void LoadCustomerProductSettings()
        {
            try
            {
                CustomerName = Microsoft.Win32.Registry.GetValue(@"HKEY_CURRENT_USER\Software\SchBomConvert",
                                                                "CustomerName", "")?.ToString() ?? "";
                ProductName = Microsoft.Win32.Registry.GetValue(@"HKEY_CURRENT_USER\Software\SchBomConvert",
                                                               "ProductName", "")?.ToString() ?? "";

                CustomerNameTextBox.Text = CustomerName;
                ProductNameTextBox.Text = ProductName;
            }
            catch (Exception)
            {
                CustomerName = "";
                ProductName = "";
            }
        }
        // 匯出按鈕點擊事件
        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            if (BomPreviewList.Count == 0)
            {
                MessageBox.Show("沒有資料可以匯出，請先載入 BOM 資料。", "匯出失敗", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            CustomerName = CustomerNameTextBox.Text?.Trim() ?? "";
            ProductName = ProductNameTextBox.Text?.Trim() ?? "";

            string originalFileName = Path.GetFileNameWithoutExtension(FilePathTextBox.Text);
            var saveFileDialog = new SaveFileDialog
            {
                Title = "匯出 BOM 資料",
                Filter = "Excel 檔案 (*.xlsx)|*.xlsx",
                DefaultExt = "xlsx",
                FileName = $"{originalFileName}_BOM_{DateTime.Now:HHmmss}.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    ExportToExcel(saveFileDialog.FileName);
                    MessageBox.Show($"匯出成功！\n請留意匯出檔案時間是否正確(時/分/秒)\n檔案已儲存至：{saveFileDialog.FileName}",
                                  "匯出完成", MessageBoxButton.OK, MessageBoxImage.Information);
                    LogUsage("匯出Excel", System.IO.Path.GetFileName(saveFileDialog.FileName));

                    if (isAutoOpenFileEnabled)
                        OpenFile(saveFileDialog.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"匯出失敗：{ex.Message}", "錯誤，請提供原檔及錯誤報告",
                                  MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        // 匯出至 Excel
        private void ExportToExcel(string filePath)
        {
            using var workbook = new XLWorkbook();
            string originalFileName = Path.GetFileNameWithoutExtension(FilePathTextBox.Text);
            string worksheetName = SanitizeWorksheetName(originalFileName);
            var worksheet = workbook.Worksheets.Add(worksheetName);

            SetColumnWidths(worksheet);

            int currentRow = 5;//BOM標頭資訊行數
            var categories = BomPreviewList.Where(item => item.IsAltLine && item.PartName?.EndsWith("料表") == true).ToList();
            bool isFirstCategory = true;

            for (int categoryIndex = 0; categoryIndex < categories.Count; categoryIndex++)
            {
                var categoryItem = categories[categoryIndex];
                string categoryName = categoryItem.PartName ?? "";

                SetCategoryHeader(worksheet, currentRow, originalFileName, categoryName, ref isFirstCategory);
                SetTableHeaders(worksheet, currentRow);
                currentRow++;

                currentRow = ProcessCategoryItems(worksheet, categoryName, currentRow, out decimal categoryTotal);
                SetCategorySubtotal(worksheet, currentRow, categoryName, categoryTotal);
                currentRow++;

                AddStaffInfoToCategory(worksheet, currentRow); //工程人員資訊行數
                currentRow += 5;

                // 料表間間隔
                if (categoryIndex < categories.Count - 1)
                {
                    AddEmptyRows(worksheet, currentRow, 2, 35, false);
                    currentRow += 2;
                }
            }

            SetGrandTotal(worksheet, currentRow - 3);
            workbook.SaveAs(filePath);
            LogUsage("匯出Excel", System.IO.Path.GetFileName(filePath));
        }

        // 設定欄位寬度
        private void SetColumnWidths(IXLWorksheet worksheet)
        {
            var widths = new[] { 3.87, 15.27, 25, 9, 6, 41.13, 6.13, 6.13, 13, 13, 8.4 };
            for (int i = 0; i < widths.Length; i++)
                worksheet.Column(i + 1).Width = widths[i];
        }

        // 設定分類標題
        private void SetCategoryHeader(IXLWorksheet worksheet, int currentRow, string fileName, string categoryName, ref bool isFirstCategory)
        {
            SetMergedCell(worksheet, $"A{currentRow - 4}:C{currentRow - 4}", "英士得科技股份有限公司", 20, 30);

            // 使用輸入的客戶名稱
            string customerText = string.IsNullOrWhiteSpace(CustomerName) ? "客戶名稱: " : $"客戶名稱: {CustomerName}";
            SetMergedCell(worksheet, $"A{currentRow - 3}:C{currentRow - 3}", customerText, 16, 24.8);

            // 使用輸入的產品名稱
            string productText = string.IsNullOrWhiteSpace(ProductName) ? "產品名稱: " : $"產品名稱: {ProductName}";
            SetMergedCell(worksheet, $"A{currentRow - 2}:D{currentRow - 2}", productText, 16, 24.8);

            SetMergedCell(worksheet, $"A{currentRow - 1}:D{currentRow - 1}", "產品編號: " + fileName, 16, 24.8);

            if (isFirstCategory && categoryName == "SMD料表")
            {
                SetCellStyle(worksheet.Cell(currentRow - 1, 6), categoryName, 14, true);
                SetMergedCell(worksheet, $"H{currentRow - 1}:I{currentRow - 1}",
                             $"DATE:{DateTime.Now:yyyy/MM/dd}", 14, 0, true);
                isFirstCategory = false;
            }
            else if (categoryName == "DIP料表" || categoryName == "其他料表")
            {
                SetCellStyle(worksheet.Cell(currentRow - 1, 6), categoryName, 14, true);
            }
        }

        // 設定合併儲存格
        private void SetMergedCell(IXLWorksheet worksheet, string range, string value, int fontSize, double height, bool centerAlign = false)
        {
            var cellRange = worksheet.Range(range);
            cellRange.Merge().Value = value;

            var style = cellRange.Style;
            style.Font.FontName = "Baskerville";
            style.Font.FontSize = fontSize;
            style.Font.Bold = true;
            style.Alignment.Horizontal = centerAlign ? XLAlignmentHorizontalValues.Center : XLAlignmentHorizontalValues.Left;
            style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            if (height > 0)
                worksheet.Row(int.Parse(range.Split(':')[0].Substring(1))).Height = height;
        }

        // 設定儲存格樣式
        private void SetCellStyle(IXLCell cell, string value, int fontSize, bool bold)
        {
            cell.Value = value;
            var style = cell.Style;
            style.Font.FontName = "Baskerville";
            style.Font.FontSize = fontSize;
            style.Font.Bold = bold;
            style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        // 設定表格標題
        private void SetTableHeaders(IXLWorksheet worksheet, int currentRow)
        {
            var headers = new[] { "N", "零件名稱", "規格", "包裝", "數量", "零件編號", "單價", "小計", "廠商", "中國廠商", "備註" };

            for (int i = 0; i < headers.Length; i++)
                worksheet.Cell(currentRow, i + 1).Value = headers[i];

            var headerRange = worksheet.Range(currentRow, 1, currentRow, 11);
            var style = headerRange.Style;
            style.Font.FontName = "Baskerville";
            style.Font.Bold = true;
            style.Fill.BackgroundColor = XLColor.White;
            style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            style.Border.InsideBorder = XLBorderStyleValues.Thin;
        }

        // 處理分類項目
        private int ProcessCategoryItems(IXLWorksheet worksheet, string categoryName, int currentRow, out decimal categoryTotal)
        {
            var (categoryItems, ncItems) = GetCategoryItems(categoryName);
            categoryTotal = 0;
            int itemNumber = 1;

            // 處理一般項目
            foreach (var item in categoryItems)
                currentRow = ProcessItemWithNumber(worksheet, item, currentRow, ref categoryTotal, ref itemNumber);

            // 處理NC項目
            if (ncItems.Count > 0)
            {
                AddEmptyRows(worksheet, currentRow, 1, 35, true);
                currentRow++;

                foreach (var ncItem in ncItems)
                    currentRow = ProcessNCItemWithNumber(worksheet, ncItem, currentRow, ref categoryTotal, ref itemNumber);
            }

            AddEmptyRows(worksheet, currentRow, 2, 35, true);
            return currentRow + 2;
        }

        // 取得分類項目
        private (List<BomPreviewItem> categoryItems, List<BomPreviewItem> ncItems) GetCategoryItems(string categoryName)
        {
            var categoryItems = new List<BomPreviewItem>();
            var ncItems = new List<BomPreviewItem>();
            bool foundCategory = false;

            foreach (var item in BomPreviewList)
            {
                if (item.IsAltLine && item.PartName == categoryName)
                {
                    foundCategory = true;
                    continue;
                }

                if (foundCategory)
                {
                    if (item.IsAltLine && item.PartName?.EndsWith("料表") == true)
                        break;

                    // 跳過所有替代料行（它們會從主項目重新產生）
                    if (item.IsAltLine)
                        continue;

                    // 只處理主項目
                    if (!item.IsAltLine)
                    {
                        if (IsNCItem(item.Spec))
                            ncItems.Add(item);
                        else
                            categoryItems.Add(item);
                    }
                }
            }

            return (categoryItems, ncItems);
        }

        // 設定分類小計
        private void SetCategorySubtotal(IXLWorksheet worksheet, int currentRow, string categoryName, decimal categoryTotal)
        {
            var subtotalCell = worksheet.Cell(currentRow, 7);
            subtotalCell.Value = $"{categoryName.Replace("料表", "")} 小計：";
            subtotalCell.Style.Font.Bold = true;
            subtotalCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

            var totalCell = worksheet.Cell(currentRow, 8);
            totalCell.Value = categoryTotal.ToString("0.##");
            var style = totalCell.Style;
            style.Font.Bold = true;
            style.Fill.BackgroundColor = XLColor.LightYellow;
            style.Border.OutsideBorder = XLBorderStyleValues.Thick;

            worksheet.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            worksheet.Row(currentRow).Height = 24.8;
        }

        // 設定總計
        private void SetGrandTotal(IXLWorksheet worksheet, int currentRow)
        {
            decimal grandTotal = BomPreviewList.Where(item => !item.IsAltLine).Sum(item => item.Subtotal);

            var labelCell = worksheet.Cell(currentRow, 7);
            labelCell.Value = "總計：";
            labelCell.Style.Font.Bold = true;
            labelCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

            var totalCell = worksheet.Cell(currentRow, 8);
            totalCell.Value = grandTotal.ToString("0.##");
            var style = totalCell.Style;
            style.Font.Bold = true;
            style.Fill.BackgroundColor = XLColor.Yellow;
            style.Border.OutsideBorder = XLBorderStyleValues.Thick;

            worksheet.Row(currentRow).Height = 24.8;
        }

        // 添加空白行
        private void AddEmptyRows(IXLWorksheet worksheet, int startRow, int count, double height, bool withBorder)
        {
            for (int i = 0; i < count; i++)
            {
                if (withBorder)
                {
                    var range = worksheet.Range(startRow + i, 1, startRow + i, 11);
                    range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                }
                worksheet.Row(startRow + i).Height = height;
            }
        }

        // 添加工程人員資料
        private void AddStaffInfoToCategory(IXLWorksheet worksheet, int startRow)
        {
            startRow--;

            // 第一行：工作人員名單
            SetStaffInfo(worksheet, startRow, 1, 2, $"研發助理：{SelectedRDAssistant}");
            SetStaffInfo(worksheet, startRow, 3, 3, $"Layout：{SelectedLayoutPerson}");
            SetStaffInfo(worksheet, startRow, 4, 5, $"線路設計：{SelectedCircuitDesigner}");

            // 第二行：助理修改日期
            SetStaffInfo(worksheet, startRow + 1, 1, 3, $"助理修改日期：{DateTime.Now:yyyy/MM/dd}");

            // 第三行：核對人員和核對日期
            SetStaffInfo(worksheet, startRow + 2, 1, 2, "核對人員：");
            SetStaffInfo(worksheet, startRow + 2, 3, 4, "核對日期：");

            // 設定行高
            for (int i = 0; i < 4; i++)
                worksheet.Row(startRow + i).Height = 24.8;
        }

        // 設定工程人員資料
        private void SetStaffInfo(IXLWorksheet worksheet, int row, int startCol, int endCol, string value)
        {
            if (startCol != endCol)
                worksheet.Range(row, startCol, row, endCol).Merge();

            var cell = worksheet.Cell(row, startCol);
            cell.Value = value;
            var style = cell.Style;
            style.Font.FontName = "Baskerville";
            style.Font.Bold = true;
            style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        // 檢查是否為NC項目
        private bool IsNCItem(string? spec) =>
            !string.IsNullOrWhiteSpace(spec) &&
            (spec.Contains("NC", StringComparison.OrdinalIgnoreCase) ||
             spec.Contains("(NC)", StringComparison.OrdinalIgnoreCase));

        // 處理一般項目
        private int ProcessItemWithNumber(IXLWorksheet worksheet, BomPreviewItem item, int currentRow, ref decimal categoryTotal, ref int itemNumber)
        {
            if (item.IsAltLine && item.PartName?.Contains("替代料") == true)
                return ProcessAlternativeItem(worksheet, item, currentRow);

            if (!item.IsAltLine)
                return ProcessRegularItem(worksheet, item, currentRow, ref categoryTotal, ref itemNumber);

            return currentRow;
        }

        // 處理替代料項目
        private int ProcessAlternativeItem(IXLWorksheet worksheet, BomPreviewItem item, int currentRow)
        {
            worksheet.Cell(currentRow, 2).Value = item.PartName;
            worksheet.Cell(currentRow, 3).Value = FormatSpecForDisplay(item.Spec);  // <-- 加上格式化

            // 設定廠商欄位
            bool isChinese = item.PartName.Contains("中國") || item.VendorCN?.Contains("中國") == true;
            int vendorCol = isChinese ? 10 : 9;
            SetVendorCellStyle(worksheet.Cell(currentRow, vendorCol), item.VendorCN ?? "");

            // 設定替代料樣式
            var altRange = worksheet.Range(currentRow, 1, currentRow, 11);
            SetAlternativeItemStyle(altRange);
            worksheet.Row(currentRow).Height = 35;

            return currentRow + 1;
        }
        private void SetVendorCellStyle(IXLCell cell, string value, bool isNC = false)
        {
            cell.Value = value;
            var style = cell.Style;
            style.Font.FontName = "Baskerville";
            style.Font.FontSize = 8;
            style.Font.Bold = true;
            style.Alignment.WrapText = true;
            style.Alignment.Vertical = XLAlignmentVerticalValues.Top;

            // 如果是NC項目，保持藍色
            if (isNC)
            {
                style.Font.FontColor = XLColor.DodgerBlue;
            }
        }

        // 設定替代料項目樣式
        private void SetAlternativeItemStyle(IXLRange range)
        {
            var style = range.Style;
            style.Font.FontName = "Baskerville";
            style.Font.Bold = true;
            style.Font.Italic = true;
            style.Font.FontColor = XLColor.FromArgb(64, 64, 64);
            style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        // 處理一般項目
        private int ProcessRegularItem(IXLWorksheet worksheet, BomPreviewItem item, int currentRow, ref decimal categoryTotal, ref int itemNumber)
        {
            string partNumbers = FormatPartNumbers(item.Code ?? "");
            var parts = SplitPartNumbers(partNumbers);
            int totalParts = parts.Count;
            int mainRowParts = Math.Min(8, totalParts);
            int additionalRowsNeeded = Math.Max(0, (totalParts - 8 + 7) / 8);

            // 填充主要行
            FillMainRow(worksheet, item, currentRow, parts.Take(mainRowParts), itemNumber.ToString());
            ApplyMainRowStyles(worksheet, currentRow);
            SetRowBorder(worksheet, currentRow);

            categoryTotal += item.Subtotal;
            itemNumber++;
            currentRow++;

            // 處理額外的零件編號行
            currentRow = ProcessAdditionalPartRows(worksheet, parts, currentRow, mainRowParts, additionalRowsNeeded);

            // 新增：處理此項目的替代料
            if (!string.IsNullOrWhiteSpace(item.Alt1))
            {
                //var (spec, vendor) = ParseAlternativePart(item.Alt1.Trim());
                var (spec, vendor, flagged) = AltPartRules.Parse(item.Alt1.Trim());
                var altItem = new BomPreviewItem
                {
                    IsAltLine = true,
                    PartName = "替代料:",
                    Spec = spec,
                    VendorCN = vendor,
                    Index = null
                };
                currentRow = ProcessAlternativeItem(worksheet, altItem, currentRow);
            }

            if (!string.IsNullOrWhiteSpace(item.Alt2))
            {
                //var (spec, vendor) = ParseAlternativePart(item.Alt2.Trim());
                var (spec, vendor, flagged) = AltPartRules.Parse(item.Alt2.Trim());
                var altItem = new BomPreviewItem
                {
                    IsAltLine = true,
                    PartName = "替代料2:",
                    Spec = spec,
                    VendorCN = vendor,
                    Index = null
                };
                currentRow = ProcessAlternativeItem(worksheet, altItem, currentRow);
            }

            return currentRow;
        }

        // 填充主要行
        private void FillMainRow(IXLWorksheet worksheet, BomPreviewItem item, int currentRow, IEnumerable<string> parts, string number)
        {
            worksheet.Cell(currentRow, 1).Value = number;
            worksheet.Cell(currentRow, 2).Value = item.PartName ?? "";
            worksheet.Cell(currentRow, 3).Value = FormatSpecForDisplay(item.Spec);  // <-- 加上格式化
            worksheet.Cell(currentRow, 4).Value = item.Package ?? "";
            worksheet.Cell(currentRow, 5).Value = item.QuantityDisplay;
            worksheet.Cell(currentRow, 6).Value = string.Join(", ", parts);
            worksheet.Cell(currentRow, 7).Value = item.UnitPrice > 0 ? item.UnitPriceDisplay : "";
            worksheet.Cell(currentRow, 8).Value = item.UnitPrice > 0 ? item.SubtotalDisplay : "";
            worksheet.Cell(currentRow, 9).Value = item.Vendor ?? "";
            worksheet.Cell(currentRow, 10).Value = item.VendorCN ?? "";
            worksheet.Cell(currentRow, 11).Value = item.Note ?? "";
        }

        // 處理額外零件編號行
        private int ProcessAdditionalPartRows(IXLWorksheet worksheet, List<string> parts, int currentRow, int mainRowParts, int additionalRowsNeeded)
        {
            int partIndex = mainRowParts;
            for (int extraRow = 0; extraRow < additionalRowsNeeded; extraRow++)
            {
                int remainingParts = parts.Count - partIndex;
                int partsInThisRow = Math.Min(8, remainingParts);
                var extraParts = parts.Skip(partIndex).Take(partsInThisRow);

                worksheet.Cell(currentRow, 6).Value = string.Join(", ", extraParts);
                ApplyExtraRowStyles(worksheet, currentRow);
                SetRowBorder(worksheet, currentRow);

                partIndex += partsInThisRow;
                currentRow++;
            }
            return currentRow;
        }

        // 處理NC項目
        private int ProcessNCItemWithNumber(IXLWorksheet worksheet, BomPreviewItem item, int currentRow, ref decimal categoryTotal, ref int itemNumber)
        {
            if (item.IsAltLine && item.PartName?.Contains("替代料") == true)
                return ProcessNCAlternativeItem(worksheet, item, currentRow);

            if (!item.IsAltLine)
                return ProcessNCRegularItem(worksheet, item, currentRow, ref categoryTotal);

            return currentRow;
        }

        // 處理NC替代料項目（完整版）
        private int ProcessNCAlternativeItem(IXLWorksheet worksheet, BomPreviewItem item, int currentRow)
        {
            worksheet.Cell(currentRow, 2).Value = item.PartName;
            worksheet.Cell(currentRow, 3).Value = FormatSpecForDisplay(item.Spec);

            bool isChinese = item.PartName.Contains("中國") || item.VendorCN?.Contains("中國") == true;
            int vendorCol = isChinese ? 10 : 9;

            // 設定廠商欄位的值和樣式
            var vendorCell = worksheet.Cell(currentRow, vendorCol);
            vendorCell.Value = item.VendorCN ?? "";
            vendorCell.Style.Font.FontName = "Baskerville";
            vendorCell.Style.Font.FontSize = 8;
            vendorCell.Style.Font.Bold = true;
            vendorCell.Style.Font.FontColor = XLColor.DodgerBlue;  // <-- 保持藍色
            vendorCell.Style.Alignment.WrapText = true;
            vendorCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;

            var altRange = worksheet.Range(currentRow, 1, currentRow, 11);
            SetNCAlternativeItemStyle(altRange);
            worksheet.Row(currentRow).Height = 34.9;

            return currentRow + 1;
        }

        // 設定NC替代料項目樣式
        private void SetNCAlternativeItemStyle(IXLRange range)
        {
            var style = range.Style;
            style.Font.FontName = "Baskerville";
            style.Font.Bold = true;
            style.Font.Italic = true;
            style.Font.FontColor = XLColor.DodgerBlue;  // <-- 改為藍色（與NC主項目相同）
            style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;  // <-- 改為靠左（與一般替代料一致）
            style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            style.Border.InsideBorder = XLBorderStyleValues.Thin;  // <-- 加上內邊框
        }

        // 處理NC一般項目
        private int ProcessNCRegularItem(IXLWorksheet worksheet, BomPreviewItem item, int currentRow, ref decimal categoryTotal)
        {
            string partNumbers = FormatPartNumbers(item.Code ?? "");
            var parts = SplitPartNumbers(partNumbers);
            int totalParts = parts.Count;
            int mainRowParts = Math.Min(8, totalParts);
            int additionalRowsNeeded = Math.Max(0, (totalParts - 8 + 7) / 8);

            // 填充主要行 (NC項目序號為"預留")
            FillMainRow(worksheet, item, currentRow, parts.Take(mainRowParts), "預留");
            ApplyMainRowStyles(worksheet, currentRow);

            // NC項目特殊樣式
            var cell = worksheet.Cell(currentRow, 1);
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Font.FontSize = 10;

            var range = worksheet.Range(currentRow, 1, currentRow, 11);
            SetNCItemStyle(range);

            // 檢查NC項目的包裝欄位並設為白色
            var packageCell = worksheet.Cell(currentRow, 4);
            string packageValue = packageCell.Value.ToString();
            string[] smdSizes = { "0402", "0603", "0805", "1206", "1210", "2010", "1812", "2512", "2728", "3920", "SR3920" };

            if (smdSizes.Any(size => packageValue.Contains(size)))
            {
                packageCell.Style.Font.FontColor = XLColor.White;
            }

            worksheet.Row(currentRow).Height = 34.9;
            categoryTotal += item.Subtotal;
            currentRow++;

            // 處理額外零件編號行
            currentRow = ProcessNCAdditionalPartRows(worksheet, parts, currentRow, mainRowParts, additionalRowsNeeded);

            // 移到這裡：處理NC項目的替代料！
            if (!string.IsNullOrWhiteSpace(item.Alt1))
            {
                //var (spec, vendor) = ParseAlternativePart(item.Alt1.Trim());
                var (spec, vendor, flagged) = AltPartRules.Parse(item.Alt1.Trim());
                var altItem = new BomPreviewItem
                {
                    IsAltLine = true,
                    PartName = "替代料:",
                    Spec = spec,
                    VendorCN = vendor,
                    Index = null
                };
                currentRow = ProcessNCAlternativeItem(worksheet, altItem, currentRow);
            }

            if (!string.IsNullOrWhiteSpace(item.Alt2))
            {
                //var (spec, vendor) = ParseAlternativePart(item.Alt2.Trim());
                var (spec, vendor, flagged) = AltPartRules.Parse(item.Alt2.Trim());
                var altItem = new BomPreviewItem
                {
                    IsAltLine = true,
                    PartName = "替代料2:",
                    Spec = spec,
                    VendorCN = vendor,
                    Index = null
                };
                currentRow = ProcessNCAlternativeItem(worksheet, altItem, currentRow);
            }

            return currentRow;  // <-- return 在最後
        }

        // 處理NC額外零件編號行
        private int ProcessNCAdditionalPartRows(IXLWorksheet worksheet, List<string> parts, int currentRow, int mainRowParts, int additionalRowsNeeded)
        {
            int partIndex = mainRowParts;
            for (int extraRow = 0; extraRow < additionalRowsNeeded; extraRow++)
            {
                int remainingParts = parts.Count - partIndex;
                int partsInThisRow = Math.Min(8, remainingParts);
                var extraParts = parts.Skip(partIndex).Take(partsInThisRow);

                worksheet.Cell(currentRow, 6).Value = string.Join(", ", extraParts);
                ApplyExtraRowStyles(worksheet, currentRow);

                var range = worksheet.Range(currentRow, 1, currentRow, 11);
                SetNCItemStyle(range);

                // 統一設置NC額外行的行高為34.9
                var row = worksheet.Row(currentRow);
                row.Height = 34.9;
                row.Style.Alignment.WrapText = true;

                partIndex += partsInThisRow;
                currentRow++;
            }
            return currentRow;
        }

        // 設定NC項目樣式
        private void SetNCItemStyle(IXLRange range)
        {
            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            range.Style.Font.FontColor = XLColor.DodgerBlue;
        }

        // 設定行邊框 - 如果這個方法也用於NC項目，需要更新行高
        private void SetRowBorder(IXLWorksheet worksheet, int currentRow)
        {
            var range = worksheet.Range(currentRow, 1, currentRow, 11);
            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            worksheet.Row(currentRow).Height = 34.9; // 統一設置為34.9 (原本是35)
        }

        // 分割零件編號
        private List<string> SplitPartNumbers(string partNumbers)
        {
            if (string.IsNullOrWhiteSpace(partNumbers))
                return new List<string>();

            return partNumbers.Split(new[] { ',', '，', ';', '；' }, StringSplitOptions.RemoveEmptyEntries)
                             .Select(p => p.Trim())
                             .Where(p => !string.IsNullOrWhiteSpace(p))
                             .ToList();
        }

        // 套用主要行樣式
        private void ApplyMainRowStyles(IXLWorksheet worksheet, int currentRow)
        {
            var columnStyles = new Dictionary<int, (int fontSize, XLAlignmentHorizontalValues horizontal, bool wrapText)>
    {
        { 3, (12, XLAlignmentHorizontalValues.Left, false) },
        { 4, (9, XLAlignmentHorizontalValues.Center, true) },
        { 6, (12, XLAlignmentHorizontalValues.Left, true) },
        { 7, (12, XLAlignmentHorizontalValues.Left, false) },
        { 8, (12, XLAlignmentHorizontalValues.Left, false) },
        { 9, (8, XLAlignmentHorizontalValues.Left, true) },
        { 10, (8, XLAlignmentHorizontalValues.Left, true) },
        { 11, (8, XLAlignmentHorizontalValues.Left, true) }
    };

            for (int col = 1; col <= 11; col++)
            {
                var cell = worksheet.Cell(currentRow, col);
                var style = cell.Style;

                style.Font.FontName = "Baskerville";
                style.Font.Bold = true;
                style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                // 修正：檢查包裝欄位(第4欄)是否包含SMD尺寸
                if (col == 4)
                {
                    string packageValue = cell.Value.ToString();
                    string[] smdSizes = { "0402", "0603", "0805", "1206", "1210", "2010", "1812", "2512", "2728", "3920", "SR3920" };

                    if (smdSizes.Any(size => packageValue.Contains(size)))
                    {
                        style.Font.FontColor = XLColor.White;
                    }
                }

                if (columnStyles.ContainsKey(col))
                {
                    var (fontSize, horizontal, wrapText) = columnStyles[col];
                    style.Font.FontSize = fontSize;
                    style.Alignment.Horizontal = horizontal;
                    style.Alignment.WrapText = wrapText;
                }
                else
                {
                    style.Font.FontSize = 12;
                    style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    style.Alignment.WrapText = true;
                }
            }
        }

        // 套用額外行樣式
        private void ApplyExtraRowStyles(IXLWorksheet worksheet, int currentRow)
        {
            var cell = worksheet.Cell(currentRow, 6);
            var style = cell.Style;
            style.Font.FontSize = 12;
            style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            style.Font.FontName = "Baskerville";
            style.Font.Bold = true;
            style.Alignment.WrapText = true;

            for (int col = 1; col <= 11; col++)
            {
                if (col != 6)
                {
                    var cellStyle = worksheet.Cell(currentRow, col).Style;
                    cellStyle.Font.FontName = "Baskerville";
                    cellStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cellStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                }
            }
        }
        // 格式化零件編號
        private string FormatPartNumbers(string input) =>
            string.IsNullOrWhiteSpace(input) ? "" : System.Text.RegularExpressions.Regex.Replace(input.Trim(), @"\s+", " ");

        private void RDAssistantComboBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is ComboBox comboBox)
            {
                SelectedRDAssistant = comboBox.Text;
                if (CurrentRDAssistantText != null)
                    CurrentRDAssistantText.Text = $"研發助理：{SelectedRDAssistant}";
                
                // 不再自動添加到列表中，而是等用戶點擊"新增"按鈕
            }
        }

        private void LayoutPersonComboBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is ComboBox comboBox)
            {
                SelectedLayoutPerson = comboBox.Text;
                if (CurrentLayoutPersonText != null)
                    CurrentLayoutPersonText.Text = $"Layout人員：{SelectedLayoutPerson}";
                
                // 不再自動添加到列表中，而是等用戶點擊"新增"按鈕
            }
        }

        private void CircuitDesignerComboBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is ComboBox comboBox)
            {
                SelectedCircuitDesigner = comboBox.Text;
                if (CurrentCircuitDesignerText != null)
                    CurrentCircuitDesignerText.Text = $"線路設計：{SelectedCircuitDesigner}";
                
                // 不再自動添加到列表中，而是等用戶點擊"新增"按鈕
            }
        }
        
        // 添加研發助理人員
        private void AddRDAssistant_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(SelectedRDAssistant))
                    return;
                    
                if (!RDAssistantList.Contains(SelectedRDAssistant))
                {
                    RDAssistantList.Add(SelectedRDAssistant);
                    MessageBox.Show($"已新增研發助理人員：{SelectedRDAssistant}", "新增成功", MessageBoxButton.OK, MessageBoxImage.Information);
                    SaveStaffSettings(); // 保存更新的列表
                }
                else
                {
                    MessageBox.Show($"研發助理人員「{SelectedRDAssistant}」已存在", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"新增研發助理人員時發生錯誤：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        // 刪除研發助理人員
        private void RemoveRDAssistant_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(SelectedRDAssistant))
                    return;
                    
                if (RDAssistantList.Contains(SelectedRDAssistant))
                {
                    if (MessageBox.Show($"確定要刪除研發助理人員「{SelectedRDAssistant}」嗎？", "確認刪除", 
                        MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                    {
                        // 移除列表中的項目
                        RDAssistantList.Remove(SelectedRDAssistant);
                        
                        // 如果列表不為空，選擇第一項
                        if (RDAssistantList.Count > 0)
                        {
                            SelectedRDAssistant = RDAssistantList[0];
                        }
                        else
                        {
                            SelectedRDAssistant = "";
                        }
                        
                        // 更新顯示
                        if (CurrentRDAssistantText != null)
                            CurrentRDAssistantText.Text = $"研發助理：{SelectedRDAssistant}";
                        
                        SaveStaffSettings(); // 保存更新的列表
                    }
                }
                else
                {
                    MessageBox.Show($"研發助理人員「{SelectedRDAssistant}」不存在於列表中", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"刪除研發助理人員時發生錯誤：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        // 添加Layout人員
        private void AddLayoutPerson_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(SelectedLayoutPerson))
                    return;
                    
                if (!LayoutPersonList.Contains(SelectedLayoutPerson))
                {
                    LayoutPersonList.Add(SelectedLayoutPerson);
                    MessageBox.Show($"已新增Layout人員：{SelectedLayoutPerson}", "新增成功", MessageBoxButton.OK, MessageBoxImage.Information);
                    SaveStaffSettings(); // 保存更新的列表
                }
                else
                {
                    MessageBox.Show($"Layout人員「{SelectedLayoutPerson}」已存在", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"新增Layout人員時發生錯誤：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        // 刪除Layout人員
        private void RemoveLayoutPerson_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(SelectedLayoutPerson))
                    return;
                    
                if (LayoutPersonList.Contains(SelectedLayoutPerson))
                {
                    if (MessageBox.Show($"確定要刪除Layout人員「{SelectedLayoutPerson}」嗎？", "確認刪除", 
                        MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                    {
                        // 移除列表中的項目
                        LayoutPersonList.Remove(SelectedLayoutPerson);
                        
                        // 如果列表不為空，選擇第一項
                        if (LayoutPersonList.Count > 0)
                        {
                            SelectedLayoutPerson = LayoutPersonList[0];
                        }
                        else
                        {
                            SelectedLayoutPerson = "";
                        }
                        
                        // 更新顯示
                        if (CurrentLayoutPersonText != null)
                            CurrentLayoutPersonText.Text = $"Layout人員：{SelectedLayoutPerson}";
                        
                        SaveStaffSettings(); // 保存更新的列表
                    }
                }
                else
                {
                    MessageBox.Show($"Layout人員「{SelectedLayoutPerson}」不存在於列表中", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"刪除Layout人員時發生錯誤：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        // 添加線路設計人員
        private void AddCircuitDesigner_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(SelectedCircuitDesigner))
                    return;
                    
                if (!CircuitDesignerList.Contains(SelectedCircuitDesigner))
                {
                    CircuitDesignerList.Add(SelectedCircuitDesigner);
                    MessageBox.Show($"已新增線路設計人員：{SelectedCircuitDesigner}", "新增成功", MessageBoxButton.OK, MessageBoxImage.Information);
                    SaveStaffSettings(); // 保存更新的列表
                }
                else
                {
                    MessageBox.Show($"線路設計人員「{SelectedCircuitDesigner}」已存在", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"新增線路設計人員時發生錯誤：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        // 刪除線路設計人員
        private void RemoveCircuitDesigner_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(SelectedCircuitDesigner))
                    return;
                    
                if (CircuitDesignerList.Contains(SelectedCircuitDesigner))
                {
                    if (MessageBox.Show($"確定要刪除線路設計人員「{SelectedCircuitDesigner}」嗎？", "確認刪除", 
                        MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                    {
                        // 移除列表中的項目
                        CircuitDesignerList.Remove(SelectedCircuitDesigner);
                        
                        // 如果列表不為空，選擇第一項
                        if (CircuitDesignerList.Count > 0)
                        {
                            SelectedCircuitDesigner = CircuitDesignerList[0];
                        }
                        else
                        {
                            SelectedCircuitDesigner = "";
                        }
                        
                        // 更新顯示
                        if (CurrentCircuitDesignerText != null)
                            CurrentCircuitDesignerText.Text = $"線路設計：{SelectedCircuitDesigner}";
                        
                        SaveStaffSettings(); // 保存更新的列表
                    }
                }
                else
                {
                    MessageBox.Show($"線路設計人員「{SelectedCircuitDesigner}」不存在於列表中", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"刪除線路設計人員時發生錯誤：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RDAssistantComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (sender is ComboBox comboBox && comboBox.SelectedItem is string selectedItem)
                {
                    SelectedRDAssistant = selectedItem;
                    if (CurrentRDAssistantText != null)
                        CurrentRDAssistantText.Text = $"研發助理：{SelectedRDAssistant}";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"RDAssistantComboBox_SelectionChanged 錯誤: {ex.Message}");
            }
        }

        private void LayoutPersonComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (sender is ComboBox comboBox && comboBox.SelectedItem is string selectedItem)
                {
                    SelectedLayoutPerson = selectedItem;
                    if (CurrentLayoutPersonText != null)
                        CurrentLayoutPersonText.Text = $"Layout人員：{SelectedLayoutPerson}";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LayoutPersonComboBox_SelectionChanged 錯誤: {ex.Message}");
            }
        }

        private void CircuitDesignerComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (sender is ComboBox comboBox && comboBox.SelectedItem is string selectedItem)
                {
                    SelectedCircuitDesigner = selectedItem;
                    if (CurrentCircuitDesignerText != null)
                        CurrentCircuitDesignerText.Text = $"線路設計：{SelectedCircuitDesigner}";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"CircuitDesignerComboBox_SelectionChanged 錯誤: {ex.Message}");
            }
        }

        // 儲存設定按鈕事件
        private void SaveSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            SaveStaffSettings();
            MessageBox.Show("設定已儲存！", "儲存完成", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        // 儲存人員設定
        private void SaveStaffSettings()
        {
            try
            {
                // 設定檔放在與主程式相同的資料夾
                string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "staff_config.json");
                
                // 建立設定物件
                var staffConfig = new
                {
                    RDAssistant = SelectedRDAssistant ?? "",
                    LayoutPerson = SelectedLayoutPerson ?? "",
                    CircuitDesigner = SelectedCircuitDesigner ?? "",
                    RDAssistantList = RDAssistantList,
                    LayoutPersonList = LayoutPersonList,
                    CircuitDesignerList = CircuitDesignerList
                };
                
                // 序列化並儲存到檔案
                string jsonConfig = JsonSerializer.Serialize(staffConfig, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(configPath, jsonConfig);
                
                // 記錄人員設定修改
                string staffInfo = $"研發助理:{SelectedRDAssistant};Layout:{SelectedLayoutPerson};線路設計:{SelectedCircuitDesigner}";
                LogUsage("修改人員設定", staffInfo);
                
                AddConversionMessage($"[INFO] 人員設定已儲存至：{configPath}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ERROR] 儲存人員設定失敗: {ex.Message}");
                AddConversionMessage($"[ERROR] 儲存人員設定失敗: {ex.Message}");
            }
        }

        // 載入人員設定
        private void LoadStaffSettings()
        {
            try
            {
                string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "staff_config.json");
                
                if (File.Exists(configPath))
                {
                    string jsonConfig = File.ReadAllText(configPath);
                    var staffConfig = JsonSerializer.Deserialize<dynamic>(jsonConfig);
                    
                    // 載入選定的人員
                    SelectedRDAssistant = staffConfig.GetProperty("RDAssistant").GetString() ?? "Fish";
                    SelectedLayoutPerson = staffConfig.GetProperty("LayoutPerson").GetString() ?? "WEI";
                    SelectedCircuitDesigner = staffConfig.GetProperty("CircuitDesigner").GetString() ?? "LSP";
                    
                    // 載入人員列表
                    var rdAssistantList = staffConfig.GetProperty("RDAssistantList");
                    var layoutPersonList = staffConfig.GetProperty("LayoutPersonList");
                    var circuitDesignerList = staffConfig.GetProperty("CircuitDesignerList");
                    
                    if (rdAssistantList.ValueKind != JsonValueKind.Null)
                    {
                        var loadedList = JsonSerializer.Deserialize<List<string>>(rdAssistantList.GetRawText());
                        if (loadedList != null && loadedList.Count > 0)
                        {
                            RDAssistantList.Clear();
                            foreach (var item in loadedList)
                                RDAssistantList.Add(item);
                        }
                    }
                    
                    if (layoutPersonList.ValueKind != JsonValueKind.Null)
                    {
                        var loadedList = JsonSerializer.Deserialize<List<string>>(layoutPersonList.GetRawText());
                        if (loadedList != null && loadedList.Count > 0)
                        {
                            LayoutPersonList.Clear();
                            foreach (var item in loadedList)
                                LayoutPersonList.Add(item);
                        }
                    }
                    
                    if (circuitDesignerList.ValueKind != JsonValueKind.Null)
                    {
                        var loadedList = JsonSerializer.Deserialize<List<string>>(circuitDesignerList.GetRawText());
                        if (loadedList != null && loadedList.Count > 0)
                        {
                            CircuitDesignerList.Clear();
                            foreach (var item in loadedList)
                                CircuitDesignerList.Add(item);
                        }
                    }
                    
                    AddConversionMessage($"[INFO] 人員設定已從檔案載入：{configPath}");
                }
                else
                {
                    // 如果檔案不存在，使用預設值
                    AddConversionMessage("[INFO] 未找到設定檔，使用預設人員設定");
                }

                // 更新顯示文字
                if (CurrentRDAssistantText != null)
                    CurrentRDAssistantText.Text = $"研發助理：{SelectedRDAssistant}";
                if (CurrentLayoutPersonText != null)
                    CurrentLayoutPersonText.Text = $"Layout人員：{SelectedLayoutPerson}";
                if (CurrentCircuitDesignerText != null)
                    CurrentCircuitDesignerText.Text = $"線路設計：{SelectedCircuitDesigner}";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ERROR] 載入人員設定失敗: {ex.Message}");
                AddConversionMessage($"[ERROR] 載入人員設定失敗: {ex.Message}");
            }
        }
        private string SanitizeWorksheetName(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return "BOM清單";

            string sanitized = fileName;
            char[] invalidChars = { '\\', '/', '?', '*', '[', ']' };

            foreach (char c in invalidChars)
            {
                sanitized = sanitized.Replace(c, '_');
            }

            if (sanitized.Length > 31)
            {
                sanitized = sanitized.Substring(0, 31);
            }

            if (string.IsNullOrWhiteSpace(sanitized))
            {
                sanitized = "BOM清單";
            }

            return sanitized;
        }

        private void LogUsage(string action, string fileName = "")
        {
            string logDir = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ".sys");
            if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
            string logPath = System.IO.Path.Combine(logDir, "sysdata.dat");
            string user = Environment.UserName;
            string machine = Environment.MachineName;
            string time = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            System.IO.File.AppendAllText(logPath, $"{action},{user},{machine},{time},{fileName}\n");
            if (File.Exists(logPath))
                File.SetAttributes(logPath, FileAttributes.Hidden);
        }

        private void AddConversionMessage(string msg)
        {
            ConversionMessages.Add(msg);
        }

        private void ConversionInfoButton_Click(object sender, RoutedEventArgs e)
        {
            ShowPanel(PanelType.ConversionInfo);
        }

        // 格式化規格顯示（尺寸、數值、精度之間雙空格）
        private static string FormatSpecForDisplay(string? spec)
        {
            if (string.IsNullOrWhiteSpace(spec)) return "";

            string cleaned = spec.Trim();

            // 先統一所有空白為單一空格
            cleaned = Regex.Replace(cleaned, @"\s+", " ");

            // 在SMD尺寸後面加雙空格
            cleaned = Regex.Replace(cleaned, @"\b(0402|0603|0805|1206|1210|2010|1812|2512|2728|3920|SR3920)\b\s*", "$1  ");

            // 在數值單位組合後面加雙空格
            cleaned = Regex.Replace(cleaned, @"\b(\d+\.?\d*[RKMΩmkMGTHz]{1,3}F?)\b\s*", "$1  ");

            // 在百分比後面加雙空格  
            cleaned = Regex.Replace(cleaned, @"\b(\d+\.?\d*%)\b\s*", "$1  ");

            // 清理結尾多餘空格
            cleaned = cleaned.TrimEnd();

            return cleaned;
        }
    }
}