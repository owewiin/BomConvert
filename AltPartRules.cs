// AltPartRules.cs — 替代料解析規則（獨立檔，沿用 Compo namespace 供 MainWindow 引用）
// 註：原始檔為 Big5 編碼，曾被編譯器當 UTF-8 處理導致全形字串損壞；
//     本檔由編譯後 DLL 反編譯之邏輯重建，並以 UTF-8 修正全形括號/冒號/廠商標籤。
using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace Compo
{
    public static class AltPartRules
    {
        // 統一：全形括號/冒號正規化、移除廠商標籤、壓縮空白
        public static string Normalize(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;
            s = s.Trim();

            // 全形 → 半形
            s = s.Replace('（', '(').Replace('）', ')')
                 .Replace('：', ':');

            // 括號內外去多餘空白
            s = Regex.Replace(s, @"\(\s+", "(");
            s = Regex.Replace(s, @"\s+\)", ")");

            // 把「(廠商: )/(供應商=)」等標籤統一成括號起始
            s = Regex.Replace(s, @"\(?\s*(?:廠商|廠牌|供應商|代理商|代理|品牌|Vendor)\s*[:=：]\s*", "(",
                              RegexOptions.IgnoreCase);

            // 壓縮多餘空白
            s = Regex.Replace(s, @"[ ]{2,}", " ");

            return s.Trim();
        }

        // 解析「規格(廠商)」→ (spec, vendor, flagged)；flagged 表格式異常、需人工確認
        public static (string spec, string vendor, bool flagged) Parse(string raw)
        {
            var text = Normalize(raw);
            if (string.IsNullOrEmpty(text)) return ("", "", false);

            // 1) 標準「規格 ( 廠商 )」
            var m = Regex.Match(text, @"^(?<spec>.+?)\s*\((?<vendor>[^()]+)\)\s*$");
            if (m.Success)
                return (SanitizeSpec(m.Groups["spec"].Value),
                        SanitizeVendor(m.Groups["vendor"].Value),
                        false);

            // 2) 只有左括號沒有右括號 → 以最後一個左括號切分
            var lastL = text.LastIndexOf('(');
            var lastR = text.LastIndexOf(')');
            if (lastL >= 0 && lastR < lastL)
            {
                var spec = text.Substring(0, lastL);
                var vendor = text.Substring(lastL + 1);
                return (SanitizeSpec(spec), SanitizeVendor(vendor), true);
            }

            // 3) 無括號 → 以分隔符切，若最後一段像廠商則視為廠商
            var tokens = Regex.Split(text, @"[\/、\|;；:：]")
                              .Select(t => t.Trim())
                              .Where(t => t.Length > 0)
                              .ToArray();
            if (tokens.Length >= 2)
            {
                var tail = tokens.Last();
                if (LooksLikeVendor(tail))
                {
                    var vendor = tail;
                    var spec = text.Substring(0, text.LastIndexOf(tail));
                    return (SanitizeSpec(spec), SanitizeVendor(vendor), true);
                }
            }

            // 4) 其餘：整串當規格、無廠商
            return (SanitizeSpec(text), "", true);
        }

        private static string SanitizeSpec(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            s = s.Trim();

            // 去除頭尾括號、標點與全形空白
            s = s.Trim('(', ')', '（', '）', '-', '/', ':', '：', '，', ';', '、', '|', '\\', '　');

            // 移除殘留的標籤（如 "… ，廠商 = "）
            s = Regex.Replace(s, @"[、,]?\s*(?:廠商|廠牌|供應商|代理商|代理|品牌|Vendor)\s*(?:[:=：].*)?$", "",
                              RegexOptions.IgnoreCase).Trim();

            return s;
        }

        private static string SanitizeVendor(string v)
        {
            if (string.IsNullOrWhiteSpace(v)) return "";
            v = v.Trim().Trim('(', ')', '（', '）');

            // 統一分隔：頓號 → '/'
            v = v.Replace('、', '/');

            // 去重複（以 '/' 分隔）
            var parts = v.Split('/').Select(p => p.Trim()).Where(p => p.Length > 0).Distinct().ToArray();
            v = string.Join("/", parts);

            // 移除殘留標籤
            v = Regex.Replace(v, @"^(?:廠商|廠牌|供應商|代理商|代理|品牌|Vendor)\s*[:=：]\s*", "",
                              RegexOptions.IgnoreCase).Trim();

            return v;
        }

        private static bool LooksLikeVendor(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            if (Regex.IsMatch(s, @"[一-龥]")) return true;             // 含中文
            if (Regex.IsMatch(s, @"^[A-Z]{2,10}$")) return true;              // 2~10 大寫英文
            if (Regex.IsMatch(s, @"\b(SAMSUNG|SANYO|TAIYO|MURATA|PANASONIC|LT|TI|ON|JST|JAE|MOLEX|TE|JTC|BC|HCB|VISHAY|YAGEO|NEXPERIA|INFINEON|RENESAS|ROHM)\b",
                              RegexOptions.IgnoreCase))
                return true;
            return false;
        }
    }
}
