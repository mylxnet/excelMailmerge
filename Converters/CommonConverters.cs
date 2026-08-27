using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;
using ExcelMailMerge.Models;

namespace ExcelMailMerge.Converters;

/// <summary>
/// List是否包含某值 → bool（命名列的CheckBox双向绑定辅助）
/// </summary>
public class ListContainsConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is not System.Collections.IList list || parameter == null) return false;
        var str = parameter.ToString();
        foreach (var item in list) if (item?.ToString() == str) return true;
        return false;
    }
    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => Binding.DoNothing; // 实际勾选通过Command实现，不走ConvertBack
}

/// <summary>
/// bool → 已用/未用 图标文字
/// </summary>
public class UsedIconConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is bool b && b ? "✅已用" : "— 未用";
    }
    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => Binding.DoNothing;
}

/// <summary>
/// bool → 颜色（绿/橙）
/// </summary>
public class BoolToColorConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is bool b && b
            ? new SolidColorBrush(Color.FromRgb(0x05, 0x96, 0x69))  // 绿色
            : new SolidColorBrush(Color.FromRgb(0xD9, 0x77, 0x06)); // 橙色
    }
    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => Binding.DoNothing;
}

/// <summary>
/// bool → Visibility
/// </summary>
public class BoolToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        var b = value is bool v && v;
        if (parameter != null && parameter.ToString() == "!") b = !b;
        return b ? Visibility.Visible : Visibility.Collapsed;
    }
    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => Binding.DoNothing;
}

/// <summary>
/// Enum → 是否等于某值（用于RadioButton绑定Enum）
/// </summary>
public class EnumEqualityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value == null || parameter == null) return false;
        try
        {
            var val = System.Convert.ToInt32(value);
            var param = System.Convert.ToInt32(parameter);
            return val == param;
        }
        catch { return false; }
    }
    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is bool b && b && parameter != null) return parameter;
        return Binding.DoNothing;
    }
}

/// <summary>
/// WizardStep → 是否等于某值（向导步骤可见性/状态判断）
/// </summary>
public class StepEqualityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is WizardStep cur && parameter is WizardStep target) return cur == target;
        if (value == null || parameter == null) return false;
        try
        {
            var val = System.Convert.ToInt32(value);
            var param = System.Convert.ToInt32(parameter);
            return val == param;
        }
        catch { return false; }
    }
    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => Binding.DoNothing;
}

/// <summary>
/// WizardStep → Visibility（当前步骤等于参数时可见，否则折叠）
/// </summary>
public class StepToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        var eq = value is WizardStep cur && parameter is WizardStep target && cur == target;
        return eq ? Visibility.Visible : Visibility.Collapsed;
    }
    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => Binding.DoNothing;
}

/// <summary>
/// StepStatus → 颜色画刷（Done绿 / Current琥珀 / Pending灰）
/// </summary>
public class StepStatusToBrushConverter : IValueConverter
{
    private static readonly SolidColorBrush DoneBrush = new(Color.FromRgb(0x10, 0xB9, 0x81));
    private static readonly SolidColorBrush CurrentBrush = new(Color.FromRgb(0xD4, 0xA2, 0x56));
    private static readonly SolidColorBrush PendingBrush = new(Color.FromRgb(0xD0, 0xD5, 0xDD));

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is StepStatus s ? s switch
        {
            StepStatus.Done => DoneBrush,
            StepStatus.Current => CurrentBrush,
            _ => PendingBrush
        } : PendingBrush;
    }
    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => Binding.DoNothing;
}

/// <summary>
/// StepStatus → 显示文字（Done显示✓，其他显示序号）
/// </summary>
public class StepStatusToTextConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is StepStatus s && s == StepStatus.Done ? "✓" : parameter?.ToString() ?? "";
    }
    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => Binding.DoNothing;
}

/// <summary>
/// bool → Visibility（反向：!参数时折叠）。与 BoolToVisibilityConverter 配合"!"参数已支持
/// 这里保留独立转换器以兼容旧绑定
/// </summary>
public class InverseBoolToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is bool b && b ? Visibility.Collapsed : Visibility.Visible;
    }
    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => Binding.DoNothing;
}
