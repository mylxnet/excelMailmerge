using System.IO;
using ExcelMailMerge.Models;

namespace ExcelMailMerge.Services;

/// <summary>
/// 命名服务：根据用户选择的多列+分隔符，生成文件名/Sheet名前缀
/// </summary>
public class NamingService
{
    // Excel Sheet名最大31字符
    public const int MaxSheetNameLength = 31;
    // Windows文件名建议不超过200字符（含后缀留余地）
    public const int MaxFileNameLength = 200;
    // Windows文件名非法字符
    private static readonly char[] InvalidFileNameChars =
        Path.GetInvalidFileNameChars().Union(new[] { '/', '\\', ':', '*', '?', '"', '<', '>', '|' }).Distinct().ToArray();

    /// <summary>
    /// 为某行数据生成命名组合（原始值，未做长度截断/非法字符替换）
    /// </summary>
    public string BuildRawName(DataRow row, List<string> namingColumns, string separator)
    {
        if (namingColumns == null || namingColumns.Count == 0) return string.Empty;
        var parts = new List<string?>();
        foreach (var col in namingColumns)
        {
            if (row.Values.TryGetValue(col, out var val) && val != null)
            {
                var s = val.ToString();
                if (!string.IsNullOrWhiteSpace(s)) parts.Add(s.Trim());
            }
        }
        return string.Join(separator ?? "_", parts.Where(p => !string.IsNullOrEmpty(p)));
    }

    /// <summary>
    /// 校验所有数据行的命名组合，返回：有效行结果 + 空值行 + 重名行
    /// </summary>
    public (List<NamingResult> perRowResults, List<int> emptyNameRowIndexes, List<IGrouping<string, int>> duplicateGroups)
        ValidateAllNames(ParsedDataSource dataSource, List<string> namingColumns, string separator)
    {
        var perRow = new List<NamingResult>();
        var empties = new List<int>();
        var nameToRows = new Dictionary<string, List<int>>(StringComparer.Ordinal);

        foreach (var row in dataSource.Rows)
        {
            var raw = BuildRawName(row, namingColumns, separator);
            if (string.IsNullOrWhiteSpace(raw))
            {
                perRow.Add(new NamingResult { IsValid = false, ErrorMessage = "命名组合为空（所有命名列均无值）" });
                empties.Add(row.OriginalRowIndex);
                continue;
            }
            var sanitized = SanitizeForFileName(raw);
            perRow.Add(new NamingResult { IsValid = true, Name = sanitized });
            if (!nameToRows.ContainsKey(sanitized)) nameToRows[sanitized] = new List<int>();
            nameToRows[sanitized].Add(row.OriginalRowIndex);
        }

        var dups = nameToRows.Where(kv => kv.Value.Count > 1)
            .Select(kv => new NameDuplicateGroup { Key = kv.Key, RowIndexes = kv.Value })
            .ToList();
        // 转成IGrouping兼容形式（用匿名类托一下）
        var dupGroups = dups.GroupBy(g => g.Key, g => g.RowIndexes)
            .Select(g => (IGrouping<string, int>)new DupGrouping(g.Key, g.SelectMany(x => x))).ToList();
        return (perRow, empties, dupGroups);
    }

    /// <summary>
    /// 移除非法字符、替换为下划线
    /// </summary>
    public string SanitizeForFileName(string rawName)
    {
        if (string.IsNullOrEmpty(rawName)) return string.Empty;
        var arr = rawName.ToCharArray();
        for (int i = 0; i < arr.Length; i++)
        {
            if (Array.IndexOf(InvalidFileNameChars, arr[i]) >= 0) arr[i] = '_';
        }
        return new string(arr).Trim();
    }

    /// <summary>
    /// 截断到指定最大长度（留余地给后缀）
    /// </summary>
    public string Truncate(string name, int maxLength, int reserveForSuffix = 3)
    {
        if (string.IsNullOrEmpty(name)) return name;
        if (name.Length <= maxLength) return name;
        var safeLen = Math.Max(0, maxLength - reserveForSuffix);
        return name.Substring(0, safeLen);
    }

    /// <summary>
    /// 生成带唯一后缀的文件名（永不覆盖）
    /// </summary>
    public string EnsureUniqueFileName(string folder, string baseName, string extension)
    {
        var ext = extension.StartsWith('.') ? extension : "." + extension;
        var candidate = Path.Combine(folder, baseName + ext);
        int i = 2;
        while (File.Exists(candidate))
        {
            candidate = Path.Combine(folder, $"{baseName}_{i}{ext}");
            i++;
        }
        return candidate;
    }

    /// <summary>
    /// 生成唯一 Sheet 名：严格按用户选的命名规则（baseName），冲突时加 _2 _3 后缀
    /// 不再自动追加模板 Sheet 名作为后缀（用户明确要求）
    /// </summary>
    public string EnsureUniqueSheetName(HashSet<string> usedNames, string baseName, string originalSheetName)
    {
        // 如果模板本身有多个 Sheet，且 baseName 相同 → 需要靠 originalSheetName 区分
        // 否则直接用 baseName，冲突才加数字后缀
        var needSheetDistinction = /* 占位：后续可接设置开关 */ false;
        var candidate = needSheetDistinction
            ? (string.IsNullOrEmpty(baseName) ? originalSheetName : $"{baseName}_{originalSheetName}")
            : baseName;
        candidate = Truncate(candidate, MaxSheetNameLength);
        int i = 2;
        while (string.IsNullOrWhiteSpace(candidate) || usedNames.Contains(candidate, StringComparer.OrdinalIgnoreCase))
        {
            candidate = Truncate(baseName, MaxSheetNameLength - 2) + $"_{i}";
            i++;
        }
        usedNames.Add(candidate);
        return candidate;
    }

    private class NameDuplicateGroup { public string Key { get; set; } = string.Empty; public List<int> RowIndexes { get; set; } = new(); }

    private class DupGrouping : IGrouping<string, int>
    {
        public DupGrouping(string key, IEnumerable<int> values) { Key = key; _values = values.ToList(); }
        public string Key { get; }
        private readonly List<int> _values;
        public IEnumerator<int> GetEnumerator() => _values.GetEnumerator();
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
