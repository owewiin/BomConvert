using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;

namespace SchBom_Convert
{
    public class ComponentMapping
    {
        [JsonPropertyName("name")]
        public string Name { get; set; } = "";

        [JsonPropertyName("category")]
        public string Category { get; set; } = "";

        [JsonPropertyName("keywords")]
        public List<string> Keywords { get; set; } = new();
    }

    public sealed class ComponentMappingService
    {
        private static readonly Lazy<ComponentMappingService> _instance =
            new(() => new ComponentMappingService());

        public static ComponentMappingService Instance => _instance.Value;

        public string ConfigPath { get; }
        public List<ComponentMapping> Mappings { get; private set; } = new();

        public event EventHandler? MappingsChanged;

        private FileSystemWatcher? _watcher;
        private DateTime _lastLoad = DateTime.MinValue;

        private ComponentMappingService()
        {
            ConfigPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "component_mappings.json");
            Load();
            StartWatching();
        }

        public void Load()
        {
            try
            {
                if (!File.Exists(ConfigPath))
                {
                    Mappings = GetDefaults();
                    SaveInternal();
                    return;
                }

                var json = File.ReadAllText(ConfigPath);
                var loaded = JsonSerializer.Deserialize<List<ComponentMapping>>(json,
                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

                Mappings = (loaded == null || loaded.Count == 0) ? GetDefaults() : loaded;
                _lastLoad = DateTime.UtcNow;
            }
            catch
            {
                Mappings = GetDefaults();
            }
        }

        public void Save(List<ComponentMapping> mappings)
        {
            Mappings = mappings ?? new List<ComponentMapping>();
            SaveInternal();
            MappingsChanged?.Invoke(this, EventArgs.Empty);
        }

        public void ResetToDefaults()
        {
            Mappings = GetDefaults();
            SaveInternal();
            MappingsChanged?.Invoke(this, EventArgs.Empty);
        }

        public void ImportFrom(string path)
        {
            var json = File.ReadAllText(path);
            var loaded = JsonSerializer.Deserialize<List<ComponentMapping>>(json,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
                ?? throw new InvalidDataException("JSON 內容不是有效的元件映射陣列");
            Save(loaded);
        }

        public void ExportTo(string path)
        {
            File.WriteAllText(path, Serialize(Mappings));
        }

        private void SaveInternal()
        {
            if (_watcher != null) _watcher.EnableRaisingEvents = false;
            try
            {
                File.WriteAllText(ConfigPath, Serialize(Mappings));
            }
            finally
            {
                if (_watcher != null) _watcher.EnableRaisingEvents = true;
            }
        }

        private static string Serialize(List<ComponentMapping> mappings)
        {
            return JsonSerializer.Serialize(mappings, new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            });
        }

        private void StartWatching()
        {
            try
            {
                var dir = Path.GetDirectoryName(ConfigPath)!;
                _watcher = new FileSystemWatcher(dir, Path.GetFileName(ConfigPath))
                {
                    NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size,
                    EnableRaisingEvents = true
                };
                _watcher.Changed += OnFileChanged;
            }
            catch
            {
                _watcher = null;
            }
        }

        private void OnFileChanged(object sender, FileSystemEventArgs e)
        {
            // FileSystemWatcher 常在儲存時連觸發兩次，這裡做簡單去抖
            Thread.Sleep(200);
            try
            {
                var info = new FileInfo(ConfigPath);
                if (info.LastWriteTimeUtc <= _lastLoad) return;
            }
            catch { }

            Load();
            MappingsChanged?.Invoke(this, EventArgs.Empty);
        }

        public static List<ComponentMapping> GetDefaults()
        {
            return new List<ComponentMapping>
            {
                // SMD 細分類
                new() { Name = "SMD 合金電阻", Category = "SMD", Keywords = new() { "合金電阻" } },
                new() { Name = "SMD 電阻",   Category = "SMD", Keywords = new() { "電阻", "resistor", "res", "RES", "阻" } },
                new() { Name = "SMD 鉭電",   Category = "SMD", Keywords = new() { "鉭電", "鉭質電容" } },
                new() { Name = "SMD 電容",   Category = "SMD", Keywords = new() { "電容", "capacitor", "cap", "CAP" } },
                new() { Name = "SMD 電感",   Category = "SMD", Keywords = new() { "電感", "inductor", "ind", "IND", "感" } },
                new() { Name = "SMD 三端濾波器", Category = "SMD", Keywords = new() { "三端濾波器" } },
                new() { Name = "SMD 二極體", Category = "SMD", Keywords = new() { "一般二極體", "快速二極體", "DIODE", "二極體" } },
                new() { Name = "SMD 蕭特基", Category = "SMD", Keywords = new() { "蕭特基", "蕭特基二極體" } },
                new() { Name = "SMD 稽納",   Category = "SMD", Keywords = new() { "稽納", "ZENER", "稽鈉", "ZD" } },
                new() { Name = "SMD LED",    Category = "SMD", Keywords = new() { "led", "LED", "發光" } },
                new() { Name = "SMD 電晶體", Category = "SMD", Keywords = new() { "bjt", "BJT", "電晶體" } },
                new() { Name = "SMD MOSFET", Category = "SMD", Keywords = new() { "mos", "MOS" } },
                new() { Name = "SMD 模塊",   Category = "SMD", Keywords = new() { "模塊", "模組" } },
                new() { Name = "SMD 保險絲", Category = "SMD", Keywords = new() { "fuse", "FUSE", "保險絲" } },
                new() { Name = "SMD BEAD",   Category = "SMD", Keywords = new() { "bead", "BEAD" } },
                new() { Name = "SMD TVS",    Category = "SMD", Keywords = new() { "tvs", "TVS" } },
                new() { Name = "SMD ESD",    Category = "SMD", Keywords = new() { "esd", "ESD" } },
                new() { Name = "SMD MCU",    Category = "SMD", Keywords = new() { "mcu", "MCU" } },
                new() { Name = "SMD OP",     Category = "SMD", Keywords = new() { "op", "OP" } },
                new() { Name = "SMD 光耦合", Category = "SMD", Keywords = new() { "光耦合" } },
                new() { Name = "SMD IR",     Category = "SMD", Keywords = new() { "ir", "IR" } },
                new() { Name = "SMD 邏輯閘", Category = "SMD", Keywords = new() { "邏輯閘" } },
                new() { Name = "SMD DC/DC",  Category = "SMD", Keywords = new() { "dc/dc", "DC/DC" } },
                new() { Name = "SMD LDO",    Category = "SMD", Keywords = new() { "ldo", "LDO" } },
                new() { Name = "SMD IC",     Category = "SMD", Keywords = new() { "IC", "ISO RS485", "TOUCH IC" } },
                new() { Name = "SMD 排針",   Category = "SMD", Keywords = new() { "排針" } },
                new() { Name = "SMD 排母",   Category = "SMD", Keywords = new() { "排母" } },
                new() { Name = "SMD 端子座", Category = "SMD", Keywords = new() { "端子座", "端子" } },
                new() { Name = "SMD RJ45座", Category = "SMD", Keywords = new() { "RJ45", "RJ45座" } },
                new() { Name = "SMD 按鍵",   Category = "SMD", Keywords = new() { "按鍵", "BUTTON", "SW" } },
                new() { Name = "SMD TOUCH",  Category = "SMD", Keywords = new() { "TOUCH", "彈簧", "TOUCH 彈簧" } },
                new() { Name = "SMD NTC",    Category = "SMD", Keywords = new() { "ntc", "NTC" } },
                new() { Name = "SMD 霍爾",   Category = "SMD", Keywords = new() { "霍爾" } },
                new() { Name = "SMD 壓力感測器", Category = "SMD", Keywords = new() { "壓力" } },
                new() { Name = "SMD 電流感測器", Category = "SMD", Keywords = new() { "電流" } },
                new() { Name = "SMD 晶振",   Category = "SMD", Keywords = new() { "震盪器", "振盪器" } },
                new() { Name = "SMD 銅塊",   Category = "SMD", Keywords = new() { "銅塊" } },
                new() { Name = "SMD 跳線",   Category = "SMD", Keywords = new() { "跳線" } },
                new() { Name = "SMD 銅柱",   Category = "SMD", Keywords = new() { "銅柱" } },

                // DIP 細分類
                new() { Name = "DIP 電阻",   Category = "DIP", Keywords = new() { "dip電阻", "DIP電阻" } },
                new() { Name = "DIP 水泥電阻", Category = "DIP", Keywords = new() { "水泥電阻" } },
                new() { Name = "DIP 電解電容", Category = "DIP", Keywords = new() { "金屬皮膜電容", "電解電容", "電容" } },
                new() { Name = "DIP 安規電容", Category = "DIP", Keywords = new() { "y電容", "Y電容", "x電容", "X電容", "Y2 CAP", "X1 CAP", "X1 電容", "Y1 電容", "Y2 電容" } },
                new() { Name = "DIP 電感",   Category = "DIP", Keywords = new() { "電感" } },
                new() { Name = "DIP BEAD",   Category = "DIP", Keywords = new() { "BEAD" } },
                new() { Name = "DIP CHOKE",  Category = "DIP", Keywords = new() { "choke", "CHOKE" } },
                new() { Name = "DIP 隔離變壓器", Category = "DIP", Keywords = new() { "隔離變壓器" } },
                new() { Name = "DIP 變壓器", Category = "DIP", Keywords = new() { "變壓器", "THANSFORMER", "THANS" } },
                new() { Name = "DIP 橋式",   Category = "DIP", Keywords = new() { "BRIDGE", "橋式", "橋式整流器" } },
                new() { Name = "DIP 二極體", Category = "DIP", Keywords = new() { "二極體", "diode", "DIODE" } },
                new() { Name = "DIP 蕭特基", Category = "DIP", Keywords = new() { "蕭特基", "蕭特基二極體" } },
                new() { Name = "DIP 稽納",   Category = "DIP", Keywords = new() { "稽鈉", "ZENER", "稽納" } },
                new() { Name = "DIP LED",    Category = "DIP", Keywords = new() { "led", "LED", "發光二極體" } },
                new() { Name = "DIP 保險絲", Category = "DIP", Keywords = new() { "保險絲", "FUSE" } },
                new() { Name = "DIP 突波",   Category = "DIP", Keywords = new() { "突波" } },
                new() { Name = "DIP NTC",    Category = "DIP", Keywords = new() { "NTC" } },
                new() { Name = "DIP 雷擊保護", Category = "DIP", Keywords = new() { "雷擊保護器" } },
                new() { Name = "DIP BJT",    Category = "DIP", Keywords = new() { "BJT" } },
                new() { Name = "DIP MOSFET", Category = "DIP", Keywords = new() { "MOS", "MOSFET" } },
                new() { Name = "DIP IGBT",   Category = "DIP", Keywords = new() { "IGBT" } },
                new() { Name = "DIP RELAY",  Category = "DIP", Keywords = new() { "relay", "RELAY", "繼電器" } },
                new() { Name = "DIP MCU",    Category = "DIP", Keywords = new() { "MCU" } },
                new() { Name = "DIP OP",     Category = "DIP", Keywords = new() { "OP" } },
                new() { Name = "DIP IPM",    Category = "DIP", Keywords = new() { "IPM" } },
                new() { Name = "DIP 模組",   Category = "DIP", Keywords = new() { "模組", "模塊" } },
                new() { Name = "DIP 感測器", Category = "DIP", Keywords = new() { "感測器", "SENSOR" } },
                new() { Name = "DIP 編碼開關", Category = "DIP", Keywords = new() { "編碼", "編碼開關" } },
                new() { Name = "DIP 開關",   Category = "DIP", Keywords = new() { "開關" } },
                new() { Name = "DIP 按鍵",   Category = "DIP", Keywords = new() { "按鍵", "BUTTON", "SW", "button" } },
                new() { Name = "DIP TOUCH",  Category = "DIP", Keywords = new() { "TOUCH", "彈簧", "TOUCH 彈簧" } },
                new() { Name = "DIP BUZZER", Category = "DIP", Keywords = new() { "蜂鳴器", "BUZZER" } },
                new() { Name = "DIP DISPLAY", Category = "DIP", Keywords = new() { "顯示器", "DISPLAY", "DSP" } },
                new() { Name = "DIP 排針",   Category = "DIP", Keywords = new() { "排針" } },
                new() { Name = "DIP 排母",   Category = "DIP", Keywords = new() { "排母" } },
                new() { Name = "DIP 歐規端子座", Category = "DIP", Keywords = new() { "歐規", "歐規端子", "歐規母座", "歐規公座" } },
                new() { Name = "DIP 端子座", Category = "DIP", Keywords = new() { "端子座", "端子" } },
                new() { Name = "DIP RJ45座", Category = "DIP", Keywords = new() { "RJ45", "RJ45座" } },
                new() { Name = "DIP 散熱片", Category = "DIP", Keywords = new() { "散熱片", "HEATSINK" } },
                new() { Name = "DIP 銅塊",   Category = "DIP", Keywords = new() { "銅塊" } },
                new() { Name = "DIP 銅柱",   Category = "DIP", Keywords = new() { "銅柱" } },
                new() { Name = "DIP 銅排",   Category = "DIP", Keywords = new() { "銅排" } },
            };
        }
    }
}
