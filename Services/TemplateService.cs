using System.Text.RegularExpressions;
using ClosedXML.Excel;
using ExcelMailMerge.Models;

namespace ExcelMailMerge.Services;

/// <summary>
/// 模板扫描服务：扫描模板所有占位符、匹配数据源列
/// </summary>
public class TemplateService
{
    // 占位符正则：非贪婪匹配 {...}，支持同一单元格多个
    private static readonly Regex PlaceholderRegex = new(@"\{(.+?)\}", RegexOptions.Compiled);

    /// <summary>
    /// 获取模板所有Sheet名
    /// </summary>
    public List<string> GetSheetNames(string templatePath)
    {
        using var wb = new XLWorkbook(templatePath);
        return wb.Worksheets.Select(ws => ws.Name).ToList();
    }

    /// <summary>
    /// 读取模板内容预览：返回模板中所有单元格的原始文本内容
    /// 列名为 Excel 列标（A、B、C...），便于展示多列布局
    /// </summary>
    public (List<TemplateContentRow> rows, List<string> columns) GetTemplateContent(string templatePath)
    {
        var rows = new List<TemplateContentRow>();
        var colSet = new HashSet<string>(StringComparer.Ordinal);

        using var wb = new XLWorkbook(templatePath);
        foreach (var ws in wb.Worksheets)
        {
            var used = ws.RangeUsed();
            if (used == null) continue;

            int firstRow = used.FirstCell().Address.RowNumber;
            int lastRow = used.LastCell().Address.RowNumber;
            int firstCol = used.FirstCell().Address.ColumnNumber;
            int lastCol = used.LastCell().Address.ColumnNumber;

            // 收集列名（Excel列标）
            for (int c = firstCol; c <= lastCol; c++)
                colSet.Add(GetExcelColumnName(c));

            // 按行读取
            for (int r = firstRow; r <= lastRow; r++)
            {
                var row = new TemplateContentRow
                {
                    SheetName = ws.Name,
                    RowIndex = r
                };
                for (int c = firstCol; c <= lastCol; c++)
                {
                    var cell = ws.Cell(r, c);
                    string text = GetCellDisplayText(cell);
                    if (!string.IsNullOrEmpty(text))
                        row.Values[GetExcelColumnName(c)] = text;
                }
                if (row.Values.Count > 0)
                    rows.Add(row);
            }
        }

        var columns = colSet.OrderBy(c => c, new ExcelColumnComparer()).ToList();
        return (rows, columns);
    }

    /// <summary>
    /// 将 Excel 列序号（1-based）转换为列标（1→A, 2→B, ..., 26→Z, 27→AA...）
    /// </summary>
    private static string GetExcelColumnName(int columnNumber)
    {
        var result = string.Empty;
        while (columnNumber > 0)
        {
            columnNumber--;
            result = (char)('A' + columnNumber % 26) + result;
            columnNumber /= 26;
        }
        return result;
    }

    /// <summary>
    /// 按 Excel 列标顺序比较（A < B < ... < Z < AA < ...）
    /// </summary>
    private class ExcelColumnComparer : IComparer<string>
    {
        public int Compare(string? a, string? b)
        {
            if (a == b) return 0;
            if (a == null) return -1;
            if (b == null) return 1;
            if (a.Length != b.Length) return a.Length.CompareTo(b.Length);
            return string.Compare(a, b, StringComparison.Ordinal);
        }
    }

    /// <summary>
    /// 获取单元格的显示文本（公式→缓存值，否则直接取值）
    /// </summary>
    private static string GetCellDisplayText(IXLCell cell)
    {
        try
        {
            if (cell.HasFormula)
            {
                var v = cell.CachedValue;
                return v.IsBlank ? string.Empty : v.ToString() ?? string.Empty;
            }
            return cell.GetString() ?? string.Empty;
        }
        catch { return string.Empty; }
    }

    /// <summary>
    /// 扫描模板，提取所有占位符，并与数据源列名做匹配
    /// </summary>
    public ScannedTemplate Scan(string templatePath, ParsedDataSource? dataSource = null)
    {
        var result = new ScannedTemplate { FilePath = templatePath };
        using var wb = new XLWorkbook(templatePath);

        // 1. 收集Sheet名
        result.SheetNames = wb.Worksheets.Select(ws => ws.Name).ToList();

        // 2. 遍历所有Sheet，所有单元格，提取占位符
        var occurrenceDict = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var ws in wb.Worksheets)
        {
            var used = ws.RangeUsed();
            if (used == null) continue;
            foreach (var cell in used.Cells())
            {
                string? content = GetCellContentForScan(cell);
                if (string.IsNullOrEmpty(content)) continue;
                var matches = PlaceholderRegex.Matches(content);
                foreach (Match m in matches)
                {
                    var key = m.Groups[1].Value.Trim(); // Trim后匹配
                    if (string.IsNullOrEmpty(key)) continue;
                    if (!occurrenceDict.ContainsKey(key)) occurrenceDict[key] = 0;
                    occurrenceDict[key]++;
                }
            }
        }

        // 3. 构造占位符列表，并与数据源匹配
        var nameSet = dataSource?.DisplayNameSet ?? new HashSet<string>();
        foreach (var kv in occurrenceDict.OrderBy(x => x.Key))
        {
            result.Placeholders.Add(new TemplatePlaceholder
            {
                Key = kv.Key,
                OccurrenceCount = kv.Value,
                IsMatched = nameSet.Contains(kv.Key)
            });
        }

        // 4. 同步回数据源列的IsUsedInTemplate
        if (dataSource != null)
        {
            var usedKeys = new HashSet<string>(occurrenceDict.Keys, StringComparer.Ordinal);
            foreach (var col in dataSource.Columns)
            {
                col.IsUsedInTemplate = usedKeys.Contains(col.DisplayName);
            }
        }

        return result;
    }

    /// <summary>
    /// 读取单元格用于扫描占位符的内容：公式→取CachedValue（已计算值）
    /// </summary>
    private static string? GetCellContentForScan(IXLCell cell)
    {
        if (cell.HasFormula)
        {
            try { var v = cell.CachedValue; return string.IsNullOrEmpty(v.ToString()) ? null : v.ToString(); }
            catch { return null; }
        }
        return cell.GetString();
    }

    /// <summary>
    /// 对单个字符串执行占位符替换（公共方法，供生成引擎调用）
    /// </summary>
    /// <param name="original">原始字符串（含{列名}）</param>
    /// <param name="row">数据行（列名→值）</param>
    /// <param name="settings">写入设置</param>
    /// <param name="outputType">输出：0=纯字符串文本 1=尝试保留原始类型object</param>
    public static (object? value, bool hasPlaceholder) ReplacePlaceholders(
        string original,
        DataRow row,
        AdvancedWriteSettings settings,
        int outputType = 0)
    {
        if (string.IsNullOrEmpty(original)) return (original, false);

        var matches = PlaceholderRegex.Matches(original);
        if (matches.Count == 0) return (original, false);

        // 情况A：整个单元格就是一个单独占位符（如整单元格="{姓名}"）→ 可以尝试保留原始数据类型
        if (matches.Count == 1 && matches[0].Index == 0 && matches[0].Length == original.Length)
        {
            var key = matches[0].Groups[1].Value.Trim();
            if (row.Values.TryGetValue(key, out var rawVal))
            {
                if (outputType == 1)
                    return (CoerceValue(rawVal, settings), true);
                return (ValueToString(rawVal, settings), true);
            }
            // 无匹配列 → 留空
            return (settings.WriteDbNullForEmpty ? DBNull.Value : string.Empty, true);
        }

        // 情况B：多占位符混排 / 占位符与文字混杂 → 只能输出字符串
        var resultStr = PlaceholderRegex.Replace(original, m =>
        {
            var key = m.Groups[1].Value.Trim();
            if (row.Values.TryGetValue(key, out var rawVal))
            {
                return ValueToString(rawVal, settings) ?? string.Empty;
            }
            // 缺失列 → 留空
            return string.Empty;
        });

        // 如果替换后全空，按设置返回DBNull
        if (string.IsNullOrEmpty(resultStr) && settings.WriteDbNullForEmpty)
            return (DBNull.Value, true);
        return (resultStr, true);
    }

    /// <summary>
    /// 值转字符串（用于混排场景）
    /// </summary>
    private static string? ValueToString(object? val, AdvancedWriteSettings settings)
    {
        if (val == null || val is DBNull) return settings.WriteDbNullForEmpty ? null : string.Empty;
        if (val is DateTime dt)
        {
            return settings.PreserveDateTimeType ? dt.ToString("yyyy-MM-dd HH:mm:ss") : dt.ToString("yyyy-MM-dd HH:mm:ss");
        }
        if (val is double d)
        {
            // 长数字 → 字符串化（防止后续写入时变科学计数，若AutoLongNumberAsText=true写入端会单独处理）
            var str = d.ToString("G");
            if (double.TryParse(str, System.Globalization.NumberStyles.Float, null, out var parsed)
                && Math.Abs(parsed - d) < 1e-9) return str;
            return d.ToString();
        }
        if (val is decimal dec) return dec.ToString("G");
        return val.ToString();
    }

    /// <summary>
    /// 按高级设置强转值类型（用于单占位符单元格，保留DateTime/数值）
    /// </summary>
    private static object? CoerceValue(object? val, AdvancedWriteSettings settings)
    {
        if (val == null || val is DBNull)
            return settings.WriteDbNullForEmpty ? DBNull.Value : string.Empty;

        // 字符串类（包括"纯数字长字符串"）
        if (val is string str)
        {
            if (settings.AutoLongNumberAsText && IsLongPureNumber(str))
                return str; // 保持字符串，写入端强制文本格式
            if (!settings.PreserveNumericType && !settings.PreserveDateTimeType)
                return str;
            // 尝试解析成数字/日期
            if (settings.PreserveNumericType &&
                decimal.TryParse(str, System.Globalization.NumberStyles.Float, null, out var dec))
                return dec;
            if (settings.PreserveDateTimeType &&
                DateTime.TryParse(str, out var dt))
                return dt;
            return str;
        }

        // 已就是DateTime
        if (val is DateTime) return settings.PreserveDateTimeType ? val : ((DateTime)val).ToString("yyyy-MM-dd HH:mm:ss");

        // 已是数值类型
        if (val is double || val is decimal || val is int || val is long || val is float)
        {
            if (settings.PreserveNumericType) return val;
            return val.ToString();
        }

        return val;
    }

    /// <summary>
    /// 判断字符串是否是纯数字且长度>11（身份证、手机号、银行卡）
    /// </summary>
    public static bool IsLongPureNumber(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return false;
        if (s.Length <= 11) return false;
        return s.All(c => c >= '0' && c <= '9');
    }
}
