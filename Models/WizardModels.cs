using ExcelMailMerge.Helpers;

namespace ExcelMailMerge.Models;

/// <summary>
/// 向导步骤（四步递进：数据源 → 模板 → 预览 → 生成）
/// </summary>
public enum WizardStep
{
    /// <summary>① 数据源：上传并校验</summary>
    DataSource = 0,
    /// <summary>② 模板：编辑占位符</summary>
    Template = 1,
    /// <summary>③ 预览：试生成第1条确认效果</summary>
    Preview = 2,
    /// <summary>④ 生成：设置并全部生成</summary>
    Generate = 3
}

/// <summary>
/// 步骤状态（用于导航圆圈视觉表现）
/// </summary>
public enum StepStatus
{
    /// <summary>已完成（绿色✓）</summary>
    Done = 0,
    /// <summary>当前（琥珀色高亮）</summary>
    Current = 1,
    /// <summary>未激活（灰色禁用）</summary>
    Pending = 2
}

/// <summary>
/// 单个步骤导航项（用于 UI 绑定，支持属性变化通知）
/// </summary>
public class WizardStepItem : ObservableObject
{
    public WizardStep Step { get; set; }

    private StepStatus _status = StepStatus.Pending;
    public StepStatus Status { get => _status; set => SetProperty(ref _status, value); }

    private bool _canNavigate;
    public bool CanNavigate { get => _canNavigate; set => SetProperty(ref _canNavigate, value); }

    public string IndexText { get; set; } = string.Empty;
    public string Label { get; set; } = string.Empty;
    public string Desc { get; set; } = string.Empty;
}

/// <summary>
/// 试生成检查项（预览阶段展示）
/// </summary>
public class TrialCheckItem
{
    /// <summary>检查类别（填充效果/样式保留/Sheet处理）</summary>
    public string Category { get; set; } = string.Empty;

    /// <summary>检查项文本</summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>明细（如 3/3）</summary>
    public string Detail { get; set; } = string.Empty;

    /// <summary>是否通过</summary>
    public bool IsPass { get; set; }

    /// <summary>图标文字（✓ / ! / ×）</summary>
    public string IconText => IsPass ? "✓" : "!";
}

