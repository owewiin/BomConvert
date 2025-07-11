using ExcelDataReader;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Diagnostics;
using ClosedXML.Excel;

using System.Text.Json;
using System.Windows.Input;

namespace SchBom_Convert
{
    public partial class MainWindow : Window
    {
        // 新增：頁面類型枚舉
        public enum PanelType
        {
            Main,
            Settings,
            Preview,
            ConversionInfo // 新增
        }
        public MainWindow()
        {
            InitializeComponent();
            LoadAutoOpenSetting(); // 載入之前儲存的設定
            LoadStaffSettings(); // 載入人員設定
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

        // 在類別中加入這個屬性來儲存自動開啟檔案的設定
        private bool isAutoOpenFileEnabled = false;
        public string CurrentDate { get; set; }
        public ObservableCollection<BomItem> BomList { get; set; } = new();
        public ObservableCollection<BomPreviewItem> BomPreviewList { get; set; } = new();
        public string SelectedRDAssistant { get; set; } = "Fish";
        public string SelectedLayoutPerson { get; set; } = "WEI";
        public string SelectedCircuitDesigner { get; set; } = "LSP";
        
        // 新增：儲存各類人員的名單列表
        public List<string> RDAssistantList { get; set; } = new List<string> { "Peggy", "Fish" };
        public List<string> LayoutPersonList { get; set; } = new List<string> { "未定", "WEI", "Jane", "Wuct", "JDLee", "Jason" };
        public List<string> CircuitDesignerList { get; set; } = new List<string> { "未定", "LSP", "Jane", "Jason", "Kevin", "Yanchi" };

        public ObservableCollection<string> ConversionMessages { get; set; } = new();

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
        // 新增：統一的頁面切換方法
        private void ShowPanel(PanelType panelType)
        {
            // 隱藏所有頁面
            MainPanel.Visibility = Visibility.Collapsed;
            SettingsPanel.Visibility = Visibility.Collapsed;
            PreviewPanel.Visibility = Visibility.Collapsed;
            ConversionInfoPanel.Visibility = Visibility.Collapsed; // 新增

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
            }
        }
        private void OpenFile(string filePath)
        {
            try
            {
                // 重點：先檢查 CheckBox 是否被勾選
                if (AutoOpenFileCheckBox.IsChecked != true)
                {
                    // 如果沒有勾選，就不開啟檔案
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
                    FilePathTextBox.Text = openFileDialog.FileName;
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

                            rawMainList.Add(new BomPreviewItem
                            {
                                IsAltLine = false,
                                PartName = GetFinalPartName(GetSafeStringValue(row, 1, "零件名稱"), partNumber),
                                Spec = GetSafeStringValue(row, 2),                    
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
                        .OrderBy(p =>
                        {
                            if (p.Category == "SMD 零件")
                            {
                                var group = GetSortGroup(p);
                                return $"A-{group.sizeOrder:D3}-{group.typeOrder:D2}-{GetSortWeight(p)}";
                            }
                            else if (p.Category == "DIP 零件")
                            {
                                return $"B-{GetDipPartOrder(p):D2}-{p.PartName}";
                            }
                            else
                            {
                                return $"Z-{p.PartName}";
                            }
                        })
                        .ToList();

                    // 新增：合併相同零件
                    var mergedList = MergeDuplicateItems(sortedList);

                    int index = 1;
                    var grouped = mergedList  // 改用 mergedList 而不是 sortedList
                        .OrderBy(p => p.Category == "SMD 零件" ? 0 : p.Category == "DIP 零件" ? 1 : 2)
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

                    // 顯示處理結果
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
                            var (spec, vendor) = ParseAlternativePart(text.Trim());

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

                    // 解析替代料的核心方法
                    (string spec, string vendor) ParseAlternativePart(string text)
                    {
                        // 嘗試標準格式：規格(廠商)
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
                            "TI", "ADI", "MAXIM", "LINEAR", "INFINEON", "VISHAY", "ROHM", "MURATA", "TDK", "IXYS", "TKS" ,
                            "SAMSUNG", "PANASONIC", "NICHICON", "RUBYCON", "KEMET", "AVX", "YAGEO", "BOURNS", "GT(Samxon)", "Samxon" ,
                            "DIODES", "FAIRCHILD", "ON", "NEXPERIA", "MICROCHIP", "ATMEL", "CYPRESS", "ALTERA",
                            "XILINX", "LATTICE", "ANALOG", "MAXLINEAR", "BROADCOM", "MARVELL", "QUALCOMM",
                            "OSRAM", "CREE", "LUMILEDS", "NICHIA", "CITIZEN", "SHARP", "TOSHIBA", "MITSUBISHI",
                            "OMRON", "TYCO", "MOLEX", "JST", "HIROSE", "SAMTEC", "TE", "AMPHENOL", "FOXCONN",
                            "立創", "嘉立創", "韋爾", "聖邦", "思瑞浦", "芯海", "兆易", "全志", "瑞芯微", "晶豐明源",
                            "上海如韻", "矽力杰", "中穎", "華大", "敏芯", "匯頂", "卓勝微", "紫光", "海思", "展訊",
                            "CORP", "CORPORATION", "TECH", "TECHNOLOGY", "SEMICONDUCTOR", "SEMI", "ELECTRONICS", "Comchip",
                            "littlefuse",
                            "耕興", "友士", "航興", "朝欣", "松川", "九寧", "奇普仕", "新承", "緯澄", "安帝", "晟通", "功得", "偉強" ,
                            "代理", "原廠", "官方", "授權"
                        };

                        var foundVendor = vendorKeywords.FirstOrDefault(keyword =>
                            text.Contains(keyword, StringComparison.OrdinalIgnoreCase));

                        if (foundVendor != null)
                        {
                            return ExtractVendorAndSpec(text, foundVendor);
                        }

                        // 分隔符分割
                        return TrySeparatorSplit(text);
                    }

                    // 提取廠商和規格
                    (string spec, string vendor) ExtractVendorAndSpec(string text, string foundVendor)
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

                    // 嘗試分隔符分割
                    (string spec, string vendor) TrySeparatorSplit(string text)
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

                    // 輔助方法
                    bool IsPartOfSpecification(string text) =>
                        new[] { "Samxon", "Samx", "CapXon", "Rubycon", "Panasonic", "Nichicon", "MST",
            "United", "Lelon", "GT", "LZ", "KZE", "KZH", "UPW", "UPS",
            "ESR", "Low", "High", "Temp", "V", "uF", "nF", "pF" }
                        .Any(marker => text.Contains(marker, StringComparison.OrdinalIgnoreCase)) || text.Length <= 6;

                    bool IsChineseName(string text) => Regex.IsMatch(text, @"[\u4e00-\u9fa5]");

                    bool IsLikelyVendorExtension(string text) =>
                        new[] { "TECH", "TECHNOLOGY", "SEMICONDUCTOR", "ELECTRONICS", "CORP", "CORPORATION", "LTD", "LIMITED", "INC" }
                        .Any(ext => text.Equals(ext, StringComparison.OrdinalIgnoreCase));

                    string CleanSpecString(string spec)
                    {
                        string cleaned = spec.Trim(' ', '-', '/', ':', '：', '，', ';', '；', '|', '\\', '　');
                        
                        // 格式化規格：在尺寸、數值、精度之間使用雙空格
                        // 例如：0805 1K 1% -> 0805  1K  1%
                        cleaned = Regex.Replace(cleaned, @"(\b(?:0402|0603|0805|1206|1210|2010|1812|2512|2728|3920|SR3920)\b)\s+", "$1  ");
                        cleaned = Regex.Replace(cleaned, @"(\b(?:[\d\.]+(?:R|K|M|Ω|ohm|p|pF|n|nF|u|uF|μF|V|W|A|H|F|Hz|dB|MHz|GHz|kHz|Hz|mW|W|mA|A|V|mV|μA|nA|pA|mH|μH|nH|pH|mF|μF|nF|pF|mΩ|μΩ|nΩ|pΩ|mW|μW|nW|pW|mJ|μJ|nJ|pJ|mK|μK|nK|pK|mT|μT|nT|pT|mG|μG|nG|pG|mS|μS|nS|pS|mB|μB|nB|pB|mC|μC|nC|pC|mD|μD|nD|pD|mE|μE|nE|pE|mF|μF|nF|pF|mG|μG|nG|pG|mH|μH|nH|pH|mI|μI|nI|pI|mJ|μJ|nJ|pJ|mK|μK|nK|pK|mL|μL|nL|pL|mM|μM|nM|pM|mN|μN|nN|pN|mO|μO|nO|pO|mP|μP|nP|pP|mQ|μQ|nQ|pQ|mR|μR|nR|pR|mS|μS|nS|pS|mT|μT|nT|pT|mU|μU|nU|pU|mV|μV|nV|pV|mW|μW|nW|pW|mX|μX|nX|pX|mY|μY|nY|pY|mZ|μZ|nZ|pZ)\b)\s+", "$1  ");
                        cleaned = Regex.Replace(cleaned, @"(\b(?:[\d\.]+%)\b)\s+", "$1  ");
                        
                        // 清理多餘的空格，但保留雙空格
                        cleaned = Regex.Replace(cleaned, @"\s{3,}", "  ");
                        
                        return cleaned.Trim();
                    }

                    bool IsLikelyPartNumber(string text) =>
                        !string.IsNullOrWhiteSpace(text) &&
                        Regex.IsMatch(text, @"[A-Za-z]\d+|[A-Za-z]+\d+[A-Za-z]*|\d+[A-Za-z]+|[A-Za-z]+\d+[-_\.]*[A-Za-z\d]*") &&
                        !Regex.IsMatch(text, @"^[\u4e00-\u9fa5]+$") &&
                        text.Length <= 30;

                    bool IsLikelyVendorName(string text)
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

                } // 結束內部方法區塊
                finally
                {
                    reader?.Dispose();
                    dataSet?.Dispose();
                }
            } // 結束 try 區塊
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
        } // 結束整個方法
          // 新增：智能選擇包裝值
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

            // 如果都沒有特殊格式，返回第一個非空值
            return validPackages.First();
        }
        private List<BomPreviewItem> MergeDuplicateItems(List<BomPreviewItem> items)
        {
            var mergedItems = new List<BomPreviewItem>();

            // 按照零件名稱、標準化規格分組
            var groupedItems = items
                .GroupBy(item => new
                {
                    PartName = item.PartName?.Trim(),
                    NormalizedSpec = NormalizeSpec(item.Spec),  // 標準化規格
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
                        Spec = GetBestSpecValue(group.Select(g => g.Spec)), // 選擇最完整的規格
                        Package = GetBestPackageValue(group.Select(g => g.Package)),
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
                    string msg = $"[INFO] 合併零件: {firstItem.PartName} ({firstItem.Spec}) - 合併了 {group.Count()} 項，總數量: {mergedItem.Quantity}";
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

        // 輔助方法：合併零件編號
        private string MergePartNumbers(IEnumerable<string?> partNumbers)
        {
            var validNumbers = partNumbers
                .Where(pn => !string.IsNullOrWhiteSpace(pn))
                .SelectMany(pn => pn!.Split(new[] { ',', '，', ';', '；' }, StringSplitOptions.RemoveEmptyEntries))
                .Select(pn => pn.Trim())
                .Where(pn => !string.IsNullOrWhiteSpace(pn))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(pn => pn)
                .ToList();

            return string.Join(", ", validNumbers);
        }

        // 輔助方法：合併備註
        private string MergeNotes(IEnumerable<string?> notes)
        {
            var validNotes = notes
                .Where(note => !string.IsNullOrWhiteSpace(note))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            return validNotes.Count > 0 ? string.Join("; ", validNotes) : "";
        }

        // 輔助方法：合併替代料
        private string MergeAlternatives(IEnumerable<string?> alternatives)
        {
            var validAlts = alternatives
                .Where(alt => !string.IsNullOrWhiteSpace(alt))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            return validAlts.Count > 0 ? string.Join("; ", validAlts) : "";
        }
        // 決定最終的零件名稱
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
                else if (trimmedCode.StartsWith("C"))
                {
                    return "SMD 電容";
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
                // 可以繼續擴展其他規則...
            }

            return ""; // 預設值
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

                // 修正：對於包裝欄位，不進行統一處理，直接返回原始值
                if (fieldType.ToLower().Contains("package") || fieldType.Contains("包裝"))
                {
                    return value; // 直接返回原始值，不進行額外處理
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

        // 修改 UnifyMultipleValues 方法，加入欄位類型參數
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
            // 定義標準元件類型映射
            var componentMappings = new Dictionary<string, string[]>
            {
                // SMD 細分類
                ["SMD 合金電阻"] = new[] { "合金電阻" },
                ["SMD 電阻"] = new[] { "電阻", "resistor", "res", "RES", "阻" },
                ["SMD 鉭電"] = new[] { "鉭電", "鉭質電容" },
                ["SMD 電容"] = new[] { "電容", "capacitor", "cap", "CAP" },
                ["SMD 電感"] = new[] { "電感", "inductor", "ind", "IND", "感" },

                ["SMD 三端濾波器"] = new[] { "三端濾波器" },

                ["SMD 二極體"] = new[] { "一般二極體","快速二極體","DIODE","二極體" },
                ["SMD 蕭特基"] = new[] { "蕭特基", "蕭特基二極體" },
                ["SMD 稽納"] = new[] { "稽納", "ZENER", "稽鈉" },
                ["SMD LED"] = new[] { "led", "LED", "發光" },

                ["SMD 電晶體"] = new[] { "bjt", "BJT", "電晶體" },
                ["SMD MOSFET"] = new[] { "mos", "MOS" },
                ["SMD 模塊"] = new[] { "模塊", "模組" },

                ["SMD 保險絲"] = new[] { "fuse", "FUSE", "保險絲" },
                ["SMD BEAD"] = new[] { "bead", "BEAD" },
                ["SMD TVS"] = new[] { "tvs", "TVS" },
                ["SMD ESD"] = new[] { "esd", "ESD" },

                ["SMD MCU"] = new[] { "mcu", "MCU" },
                ["SMD OP"] = new[] { "op", "OP" },
                ["SMD 光耦合"] = new[] { "光耦合" },
                ["SMD IR"] = new[] { "ir", "IR" },
                ["SMD 邏輯閘"] = new[] { "邏輯閘" },
                ["SMD DC/DC"] = new[] { "dc/dc", "DC/DC" },
                ["SMD LDO"] = new[] { "ldo", "LDO" },
                ["SMD IC"] = new[] { "IC" , "ISO RS485" },

                ["SMD 排針"] = new[] { "排針" },
                ["SMD 排母"] = new[] { "排母" },
                ["SMD 端子座"] = new[] { "端子座", "端子" },
                ["SMD RJ45座"] = new[] { "RJ45" , "RJ45座" },
                ["SMD 按鍵"] = new[] { "按鍵", "BUTTON" , "SW" },

                ["SMD NTC"] = new[] { "ntc", "NTC" },
                ["SMD 霍爾"] = new[] { "霍爾" },
                ["SMD 壓力感測器"] = new[] { "壓力" },
                ["SMD 電流感測器"] = new[] { "電流" },
                ["SMD 晶振"] = new[] { "震盪器", "振盪器" },
                
                ["SMD 銅塊"] = new[] { "銅塊" },
                ["SMD 跳線"] = new[] { "跳線" },
                ["SMD 銅柱"] = new[] { "銅柱" },

                // DIP 細分類
                ["DIP 電阻"] = new[] { "dip電阻", "DIP電阻" },
                ["DIP 水泥電阻"] = new[] { "水泥電阻" },
                ["DIP 電解電容"] = new[] { "金屬皮膜電容", "電解電容", "電容" },
                ["DIP 安規電容"] = new[] { "y電容", "Y電容", "x電容", "X電容", "Y2 CAP" , "X1 CAP", "X1 電容" , "Y1 電容", "Y2 電容" },
                ["DIP 電感"] = new[] { "電感" },
                ["DIP BEAD"] = new[] { "BEAD" },
                ["DIP CHOKE"] = new[] { "choke", "CHOKE", },
                ["DIP 隔離變壓器"] = new[] { "隔離變壓器" },
                ["DIP 變壓器"] = new[] { "變壓器","THANSFORMER","THANS" },

                ["DIP 橋式"] = new[] { "BRIDGE", "橋式", "橋式整流器" },
                ["DIP 二極體"] = new[] { "二極體", "diode", "DIODE" },
                ["DIP 蕭特基"] = new[] { "蕭特基", "蕭特基二極體" },
                ["DIP 稽納"] = new[] { "稽鈉","ZENER", "稽納" },
                ["DIP LED"] = new[] { "led", "LED", "發光二極體" },

                ["DIP 保險絲"] = new[] { "保險絲", "FUSE"  },
                ["DIP 突波"] = new[] { "突波" },
                ["DIP NTC"] = new[] { "NTC" },
                ["DIP 雷擊保護"] = new[] { "雷擊保護器" },
                ["DIP BJT"] = new[] { "BJT" },
                ["DIP MOSFET"] = new[] { "MOS","MOSFET" },
                ["DIP IGBT"] = new[] { "IGBT" },
                ["DIP RELAY"] = new[] { "relay", "RELAY", "繼電器" },
                ["DIP MCU"] = new[] { "MCU" },
                ["DIP OP"] = new[] { "OP" },
                ["DIP IPM"] = new[] { "IPM" },
                ["DIP 模組"] = new[] { "模組", "模塊" },

                ["DIP 感測器"] = new[] { "感測器" ,"SENSOR"},
                ["DIP 編碼開關"] = new[] { "編碼" , "編碼開關" },
                ["DIP 開關"] = new[] { "開關" },
                ["DIP 按鍵"] = new[] { "按鍵", "button" },
                ["DIP BUZZER"] = new[] { "蜂鳴器", "BUZZER" },
                ["DIP DISPLAY"] = new[] { "顯示器", "DISPLAY" },

                ["DIP 排針"] = new[] { "排針" },
                ["DIP 排母"] = new[] { "排母" },
                ["DIP 端子座"] = new[] { "端子座", "端子" },
                ["DIP 歐規"] = new[] { "歐規", "歐規端子" , "歐規母座" },
                ["DIP RJ45座"] = new[] { "RJ45" , "RJ45座" },                
                ["DIP 按鍵"] = new[] { "按鍵", "BUTTON" , "SW" },

                ["DIP 散熱片"] = new[] { "散熱片","HEATSINK" },
                ["DIP 銅塊"] = new[] { "銅塊" },
                ["DIP 銅柱"] = new[] { "銅柱" },
                ["DIP 銅牌"] = new[] { "銅牌" }
            };
            
            // 組合所有值進行分析
            string combinedText = string.Join(" ", values).ToLower();

            // 按優先級檢查元件類型 - 先檢查更具體的類型
            foreach (var mapping in componentMappings)
            {
                string targetName = mapping.Key;
                string[] keywords = mapping.Value;

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

            // 修改：只有在沒有找到任何匹配時，才使用預設值
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

        // 輔助方法：檢查是否包含SMD尺寸指示器
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

        // 保持原有的重載方法以維持相容性
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
                name.Contains("端子座") ? GetTerminalBlockOrder(spec) :
                name.Contains("端子") ? 35 :
                name.Contains("接線座") ? 35 :

                // DIP 機構件
                name.Contains("散熱片") ? 36 :
                name.Contains("HEATSINK") ? 36 :
                name.Contains("銅塊") ? 37 :
                name.Contains("銅柱") ? 38 :
                name.Contains("銅牌") ? 39 :
                name.Contains("五金") ? 40 :
                name.Contains("螺絲") ? 40 :

                // 特殊處理：根據規格判斷
                spec.Contains("uF") ? 6 :      // 電容類
                spec.Contains("MOV") ? 14 :    // 突波吸收器
                99;  // 未分類

            return dipOrder;
        }

        // 新增：端子座編號排序方法
        private static int GetTerminalBlockOrder(string spec)
        {
            if (string.IsNullOrWhiteSpace(spec)) return 34; // 預設值

            // 提取編號（如2021, 2532, 3961等）
            var numberMatch = Regex.Match(spec, @"(\d{4})");
            if (numberMatch.Success)
            {
                string number = numberMatch.Groups[1].Value;
                
                // 按照指定順序排序：2021, 2532, 3961，其餘排在後面
                return number switch
                {
                    "2021" => 34, // 第一個
                    "2532" => 35, // 第二個
                    "3961" => 36, // 第三個
                    _ => 37       // 其他編號排在最後
                };
            }

            return 34; // 沒有編號時使用預設值
        }

        // 修改後的電阻排序方法 - 按照數字大小排序
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

            // 格式化：數值(12位，含小數) + 單位權重 + 精度標記
            return $"{numericValue:000000000000.000}-{unitWeight}-{precisionSuffix}";
        }
        // 修改後的電容排序方法
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
        // 修改GetSortWeight方法中的電阻和電容部分
        private static string GetSortWeight(BomPreviewItem p)
        {
            string name = p.PartName ?? "";
            string spec = p.Spec ?? "";

            if (name.Contains("電阻")) return GetResistorSortKey(spec);
            if (name.Contains("電容")) return GetCapacitorSortKey(spec);

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
        // 匯出按鈕點擊事件
        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            if (BomPreviewList.Count == 0)
            {
                MessageBox.Show("沒有資料可以匯出，請先載入 BOM 資料。", "匯出失敗", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

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

            int currentRow = 4;
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

                AddStaffInfoToCategory(worksheet, currentRow);
                currentRow += 4;

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
            var widths = new[] { 3.87, 15.27, 22, 9, 6, 41.13, 6.13, 6.13, 13, 13, 8.4 };
            for (int i = 0; i < widths.Length; i++)
                worksheet.Column(i + 1).Width = widths[i];
        }

        // 設定分類標題
        private void SetCategoryHeader(IXLWorksheet worksheet, int currentRow, string fileName, string categoryName, ref bool isFirstCategory)
        {
            SetMergedCell(worksheet, $"A{currentRow - 3}:C{currentRow - 3}", "英士得科技股份有限公司", 20, 30);
            SetMergedCell(worksheet, $"A{currentRow - 2}:C{currentRow - 2}", "客戶名稱:", 16, 24.8);
            SetMergedCell(worksheet, $"A{currentRow - 1}:D{currentRow - 1}", fileName, 16, 24.8);

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

                    if (!item.IsAltLine && IsNCItem(item.Spec))
                        ncItems.Add(item);
                    else
                        categoryItems.Add(item);
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
            worksheet.Cell(currentRow, 3).Value = item.Spec ?? "";

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

        // 設定廠商儲存格樣式
        private void SetVendorCellStyle(IXLCell cell, string value)
        {
            cell.Value = value;
            var style = cell.Style;
            style.Font.FontName = "Baskerville";
            style.Font.FontSize = 8;
            style.Font.Bold = true;
            style.Alignment.WrapText = true;
            style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
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
            return ProcessAdditionalPartRows(worksheet, parts, currentRow, mainRowParts, additionalRowsNeeded);
        }

        // 填充主要行
        private void FillMainRow(IXLWorksheet worksheet, BomPreviewItem item, int currentRow, IEnumerable<string> parts, string number)
        {
            worksheet.Cell(currentRow, 1).Value = number;
            worksheet.Cell(currentRow, 2).Value = item.PartName ?? "";
            worksheet.Cell(currentRow, 3).Value = item.Spec ?? "";
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

        // 處理NC替代料項目
        private int ProcessNCAlternativeItem(IXLWorksheet worksheet, BomPreviewItem item, int currentRow)
        {
            worksheet.Cell(currentRow, 2).Value = item.PartName;
            worksheet.Cell(currentRow, 3).Value = item.Spec ?? "";

            bool isChinese = item.PartName.Contains("中國") || item.VendorCN?.Contains("中國") == true;
            int vendorCol = isChinese ? 10 : 9;
            SetVendorCellStyle(worksheet.Cell(currentRow, vendorCol), item.VendorCN ?? "");

            var altRange = worksheet.Range(currentRow, 1, currentRow, 11);
            SetNCAlternativeItemStyle(altRange);
            worksheet.Row(currentRow).Height = 34.9; // 統一設置為34.9

            return currentRow + 1;
        }

        // 設定NC替代料項目樣式
        private void SetNCAlternativeItemStyle(IXLRange range)
        {
            var style = range.Style;
            style.Font.FontName = "Baskerville";
            style.Font.Bold = true;
            style.Font.Italic = true;
            style.Font.FontColor = XLColor.FromArgb(64, 64, 64);
            style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            style.Border.OutsideBorder = XLBorderStyleValues.Thin;
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

            // 統一設置NC主要行的行高為34.9
            worksheet.Row(currentRow).Height = 34.9;

            categoryTotal += item.Subtotal;
            currentRow++;

            // 處理額外零件編號行
            return ProcessNCAdditionalPartRows(worksheet, parts, currentRow, mainRowParts, additionalRowsNeeded);
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

        // 修改：LostFocus 事件處理器（當用戶輸入完成並離開焦點時觸發）
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
        
        // 新增：添加研發助理人員
        private void AddRDAssistant_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(SelectedRDAssistant))
                return;
                
            if (!RDAssistantList.Contains(SelectedRDAssistant))
            {
                RDAssistantList.Add(SelectedRDAssistant);
                RDAssistantComboBox.Items.Add(new ComboBoxItem { Content = SelectedRDAssistant });
                MessageBox.Show($"已新增研發助理人員：{SelectedRDAssistant}", "新增成功", MessageBoxButton.OK, MessageBoxImage.Information);
                SaveStaffSettings(); // 保存更新的列表
            }
            else
            {
                MessageBox.Show($"研發助理人員「{SelectedRDAssistant}」已存在", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        
        // 新增：刪除研發助理人員
        private void RemoveRDAssistant_Click(object sender, RoutedEventArgs e)
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
                    
                    // 移除ComboBox中對應的項目
                    ComboBoxItem itemToRemove = null;
                    foreach (ComboBoxItem item in RDAssistantComboBox.Items)
                    {
                        if (item.Content.ToString() == SelectedRDAssistant)
                        {
                            itemToRemove = item;
                            break;
                        }
                    }
                    
                    if (itemToRemove != null)
                        RDAssistantComboBox.Items.Remove(itemToRemove);
                    
                    // 如果列表不為空，選擇第一項
                    if (RDAssistantList.Count > 0)
                    {
                        SelectedRDAssistant = RDAssistantList[0];
                        RDAssistantComboBox.Text = SelectedRDAssistant;
                    }
                    else
                    {
                        SelectedRDAssistant = "";
                        RDAssistantComboBox.Text = "";
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
        
        // 新增：添加Layout人員
        private void AddLayoutPerson_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(SelectedLayoutPerson))
                return;
                
            if (!LayoutPersonList.Contains(SelectedLayoutPerson))
            {
                LayoutPersonList.Add(SelectedLayoutPerson);
                LayoutPersonComboBox.Items.Add(new ComboBoxItem { Content = SelectedLayoutPerson });
                MessageBox.Show($"已新增Layout人員：{SelectedLayoutPerson}", "新增成功", MessageBoxButton.OK, MessageBoxImage.Information);
                SaveStaffSettings(); // 保存更新的列表
            }
            else
            {
                MessageBox.Show($"Layout人員「{SelectedLayoutPerson}」已存在", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        
        // 新增：刪除Layout人員
        private void RemoveLayoutPerson_Click(object sender, RoutedEventArgs e)
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
                    
                    // 移除ComboBox中對應的項目
                    ComboBoxItem itemToRemove = null;
                    foreach (ComboBoxItem item in LayoutPersonComboBox.Items)
                    {
                        if (item.Content.ToString() == SelectedLayoutPerson)
                        {
                            itemToRemove = item;
                            break;
                        }
                    }
                    
                    if (itemToRemove != null)
                        LayoutPersonComboBox.Items.Remove(itemToRemove);
                    
                    // 如果列表不為空，選擇第一項
                    if (LayoutPersonList.Count > 0)
                    {
                        SelectedLayoutPerson = LayoutPersonList[0];
                        LayoutPersonComboBox.Text = SelectedLayoutPerson;
                    }
                    else
                    {
                        SelectedLayoutPerson = "";
                        LayoutPersonComboBox.Text = "";
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
        
        // 新增：添加線路設計人員
        private void AddCircuitDesigner_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(SelectedCircuitDesigner))
                return;
                
            if (!CircuitDesignerList.Contains(SelectedCircuitDesigner))
            {
                CircuitDesignerList.Add(SelectedCircuitDesigner);
                CircuitDesignerComboBox.Items.Add(new ComboBoxItem { Content = SelectedCircuitDesigner });
                MessageBox.Show($"已新增線路設計人員：{SelectedCircuitDesigner}", "新增成功", MessageBoxButton.OK, MessageBoxImage.Information);
                SaveStaffSettings(); // 保存更新的列表
            }
            else
            {
                MessageBox.Show($"線路設計人員「{SelectedCircuitDesigner}」已存在", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        
        // 新增：刪除線路設計人員
        private void RemoveCircuitDesigner_Click(object sender, RoutedEventArgs e)
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
                    
                    // 移除ComboBox中對應的項目
                    ComboBoxItem itemToRemove = null;
                    foreach (ComboBoxItem item in CircuitDesignerComboBox.Items)
                    {
                        if (item.Content.ToString() == SelectedCircuitDesigner)
                        {
                            itemToRemove = item;
                            break;
                        }
                    }
                    
                    if (itemToRemove != null)
                        CircuitDesignerComboBox.Items.Remove(itemToRemove);
                    
                    // 如果列表不為空，選擇第一項
                    if (CircuitDesignerList.Count > 0)
                    {
                        SelectedCircuitDesigner = CircuitDesignerList[0];
                        CircuitDesignerComboBox.Text = SelectedCircuitDesigner;
                    }
                    else
                    {
                        SelectedCircuitDesigner = "";
                        CircuitDesignerComboBox.Text = "";
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
        // 修改現有的 SelectionChanged 事件處理器
        private void RDAssistantComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox comboBox && comboBox.SelectedItem is ComboBoxItem selectedItem)
            {
                SelectedRDAssistant = selectedItem.Content.ToString();
                comboBox.Text = SelectedRDAssistant;
                if (CurrentRDAssistantText != null)
                    CurrentRDAssistantText.Text = $"研發助理：{SelectedRDAssistant}";
            }
        }

        private void LayoutPersonComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox comboBox && comboBox.SelectedItem is ComboBoxItem selectedItem)
            {
                SelectedLayoutPerson = selectedItem.Content.ToString();
                comboBox.Text = SelectedLayoutPerson;
                if (CurrentLayoutPersonText != null)
                    CurrentLayoutPersonText.Text = $"Layout人員：{SelectedLayoutPerson}";
            }
        }

        private void CircuitDesignerComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox comboBox && comboBox.SelectedItem is ComboBoxItem selectedItem)
            {
                SelectedCircuitDesigner = selectedItem.Content.ToString();
                comboBox.Text = SelectedCircuitDesigner;
                if (CurrentCircuitDesignerText != null)
                    CurrentCircuitDesignerText.Text = $"線路設計：{SelectedCircuitDesigner}";
            }
        }

        // 新增：儲存設定按鈕事件
        private void SaveSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            SaveStaffSettings();
            MessageBox.Show("設定已儲存！", "儲存完成", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        // 修改：儲存人員設定
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
                
                // 新增：記錄人員設定修改
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

        // 修改：載入人員設定
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
                            RDAssistantList = loadedList;
                    }
                    
                    if (layoutPersonList.ValueKind != JsonValueKind.Null)
                    {
                        var loadedList = JsonSerializer.Deserialize<List<string>>(layoutPersonList.GetRawText());
                        if (loadedList != null && loadedList.Count > 0)
                            LayoutPersonList = loadedList;
                    }
                    
                    if (circuitDesignerList.ValueKind != JsonValueKind.Null)
                    {
                        var loadedList = JsonSerializer.Deserialize<List<string>>(circuitDesignerList.GetRawText());
                        if (loadedList != null && loadedList.Count > 0)
                            CircuitDesignerList = loadedList;
                    }
                    
                    AddConversionMessage($"[INFO] 人員設定已從檔案載入：{configPath}");
                }
                else
                {
                    // 如果檔案不存在，使用預設值
                    AddConversionMessage("[INFO] 未找到設定檔，使用預設人員設定");
                }

                // 延遲設定，確保 UI 已載入
                this.Loaded += (sender, e) =>
                {
                    // 設定下拉選單項目
                    if (RDAssistantComboBox != null)
                    {
                        RDAssistantComboBox.Items.Clear();
                        foreach (var name in RDAssistantList)
                        {
                            RDAssistantComboBox.Items.Add(new ComboBoxItem { Content = name });
                        }
                        RDAssistantComboBox.Text = SelectedRDAssistant;
                    }
                    
                    if (LayoutPersonComboBox != null)
                    {
                        LayoutPersonComboBox.Items.Clear();
                        foreach (var name in LayoutPersonList)
                        {
                            LayoutPersonComboBox.Items.Add(new ComboBoxItem { Content = name });
                        }
                        LayoutPersonComboBox.Text = SelectedLayoutPerson;
                    }
                    
                    if (CircuitDesignerComboBox != null)
                    {
                        CircuitDesignerComboBox.Items.Clear();
                        foreach (var name in CircuitDesignerList)
                        {
                            CircuitDesignerComboBox.Items.Add(new ComboBoxItem { Content = name });
                        }
                        CircuitDesignerComboBox.Text = SelectedCircuitDesigner;
                    }

                    // 更新右側顯示
                    if (CurrentRDAssistantText != null) CurrentRDAssistantText.Text = $"研發助理：{SelectedRDAssistant}";
                    if (CurrentLayoutPersonText != null) CurrentLayoutPersonText.Text = $"Layout人員：{SelectedLayoutPerson}";
                    if (CurrentCircuitDesignerText != null) CurrentCircuitDesignerText.Text = $"線路設計：{SelectedCircuitDesigner}";
                };
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
            string cleaned = spec.Trim(' ', '-', '/', ':', '：', '，', ';', '；', '|', '\\', '　');
            cleaned = System.Text.RegularExpressions.Regex.Replace(cleaned, @"(\b(?:0402|0603|0805|1206|1210|2010|1812|2512|2728|3920|SR3920)\b)\s+", "$1  ");
            cleaned = System.Text.RegularExpressions.Regex.Replace(cleaned, @"(\b(?:[\d\.]+(?:R|K|M|Ω|ohm|p|pF|n|nF|u|uF|μF|V|W|A|H|F|Hz|dB|MHz|GHz|kHz|Hz|mW|W|mA|A|V|mV|μA|nA|pA|mH|μH|nH|pH|mF|μF|nF|pF|mΩ|μΩ|nΩ|pΩ|mW|μW|nW|pW|mJ|μJ|nJ|pJ|mK|μK|nK|pK|mT|μT|nT|pT|mG|μG|nG|pG|mS|μS|nS|pS|mB|μB|nB|pB|mC|μC|nC|pC|mD|μD|nD|pD|mE|μE|nE|pE|mF|μF|nF|pF|mG|μG|nG|pG|mH|μH|nH|pH|mI|μI|nI|pI|mJ|μJ|nJ|pJ|mK|μK|nK|pK|mL|μL|nL|pL|mM|μM|nM|pM|mN|μN|nN|pN|mO|μO|nO|pO|mP|μP|nP|pP|mQ|μQ|nQ|pQ|mR|μR|nR|pR|mS|μS|nS|pS|mT|μT|nT|pT|mU|μU|nU|pU|mV|μV|nV|pV|mW|μW|nW|pW|mX|μX|nX|pX|mY|μY|nY|pY|mZ|μZ|nZ|pZ)\b)\s+", "$1  ");
            cleaned = System.Text.RegularExpressions.Regex.Replace(cleaned, @"(\b(?:[\d\.]+%)\b)\s+", "$1  ");
            cleaned = System.Text.RegularExpressions.Regex.Replace(cleaned, @"\s{3,}", "  ");
            return cleaned.Trim();
        }
    }
}