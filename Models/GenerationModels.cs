using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ExcelMailMerge.Models;

/// <summary>
/// 生成进度报告
/// </summary>
public class GenerationProgress : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnChanged([CallerMemberName] string? name = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        // 计算属性也依赖这些字段，联动通知
        if (name is nameof(TotalRows) or nameof(CurrentRowIndex))
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ProgressPercent)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(StatusText)));
        }
        if (name is nameof(SuccessCount) or nameof(FailCount) or nameof(IsCompleted) or nameof(IsCanceled))
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(StatusText)));
        }
    }

    private int _totalRows; public int TotalRows { get => _totalRows; set { _totalRows = value; OnChanged(); } }
    private int _currentRowIndex; public int CurrentRowIndex { get => _currentRowIndex; set { _currentRowIndex = value; OnChanged(); } }
    private int _successCount; public int SuccessCount { get => _successCount; set { _successCount = value; OnChanged(); } }
    private int _failCount; public int FailCount { get => _failCount; set { _failCount = value; OnChanged(); } }
    private int _skipCount; public int SkipCount { get => _skipCount; set { _skipCount = value; OnChanged(); } }
    private bool _isCompleted; public bool IsCompleted { get => _isCompleted; set { _isCompleted = value; OnChanged(); } }
    private bool _isCanceled; public bool IsCanceled { get => _isCanceled; set { _isCanceled = value; OnChanged(); } }
    private string _outputFolder = string.Empty; public string OutputFolder { get => _outputFolder; set { _outputFolder = value; OnChanged(); } }
    public ObservableCollection<string> Logs { get; set; } = new();
    public List<(int rowIndex, string fileName, string error)> Failures { get; set; } = new();

    public double ProgressPercent => TotalRows == 0 ? 0 : Math.Round((double)CurrentRowIndex / TotalRows * 100, 1);

    public string StatusText
    {
        get
        {
            if (IsCanceled) return $"⛔ 已取消：成功{SuccessCount}个，处理到第{CurrentRowIndex}行";
            if (IsCompleted) return FailCount == 0
                ? $"✅ 全部完成：成功{SuccessCount}个，耗时见日志"
                : $"⚠️ 完成但有失败：成功{SuccessCount}，失败{FailCount}";
            return $"⏳ 处理中：{CurrentRowIndex}/{TotalRows}（{ProgressPercent}%） 成功{SuccessCount} 失败{FailCount}";
        }
    }
}

/// <summary>
/// 命名组合结果
/// </summary>
public class NamingResult
{
    public bool IsValid { get; set; }
    public string Name { get; set; } = string.Empty;
    public string ErrorMessage { get; set; } = string.Empty;
}

/// <summary>
/// 试生成（预览）结果
/// </summary>
public class PreviewResult
{
    /// <summary>是否成功</summary>
    public bool IsSuccess { get; set; }

    /// <summary>临时预览文件路径</summary>
    public string TempFilePath { get; set; } = string.Empty;

    /// <summary>使用的第几行数据（原始行号）</summary>
    public int SourceRowIndex { get; set; }

    /// <summary>临时文件名</summary>
    public string TempFileName { get; set; } = string.Empty;

    /// <summary>包含 Sheet 数</summary>
    public int SheetCount { get; set; }

    /// <summary>已替换占位符数</summary>
    public int PlaceholderReplacedCount { get; set; }

    /// <summary>未匹配占位符数</summary>
    public int PlaceholderUnmatchedCount { get; set; }

    /// <summary>公式转值数</summary>
    public int FormulaConvertedCount { get; set; }

    /// <summary>填充了数据的模板内容预览行（用于UI显示）</summary>
    public List<TemplateContentRow> FilledContentRows { get; set; } = new();

    /// <summary>填充了数据的模板内容预览列（Excel列标 A/B/C...）</summary>
    public List<string> FilledContentColumns { get; set; } = new();

    /// <summary>错误信息</summary>
    public string ErrorMessage { get; set; } = string.Empty;
}

