using System.Collections.Specialized;
using System.ComponentModel;
using System.IO.Pipes;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Interop;
using ExcelMailMerge.Models;
using ExcelMailMerge.ViewModels;

namespace ExcelMailMerge;

public partial class MainWindow : Window
{
    private readonly MainViewModel _vm;
    private CancellationTokenSource? _pipeCts;

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    private static extern bool IsIconic(IntPtr hWnd);
    private const int SW_RESTORE = 9;

    public MainWindow()
    {
        InitializeComponent();
        _vm = new MainViewModel();
        DataContext = _vm;
        Closing += (_, _) =>
        {
            _pipeCts?.Cancel();
            _vm.SaveSettings();
        };

        // 订阅集合变化，动态重建 DataGrid 列
        _vm.Columns.CollectionChanged += (_, _) => RebuildPreviewColumns();
        _vm.TemplateContentColumns.CollectionChanged += (_, _) => RebuildTemplateContentColumns();

        // 订阅 PreviewResult 变化，重建填充预览列
        _vm.PropertyChanged += (_, e) =>
        {
            if (e.PropertyName == nameof(MainViewModel.FilledContentRows) ||
                e.PropertyName == nameof(MainViewModel.FilledContentColumns))
            {
                RebuildFilledPreviewColumns();
            }
        };

        RebuildPreviewColumns();
        RebuildTemplateContentColumns();

        // 启动命名管道监听，接收第二个实例的激活请求
        StartPipeListener();
    }

    private void StartPipeListener()
    {
        _pipeCts = new CancellationTokenSource();
        var token = _pipeCts.Token;
        ThreadPool.QueueUserWorkItem(async _ =>
        {
            while (!token.IsCancellationRequested)
            {
                try
                {
                    using var server = new NamedPipeServerStream("ExcelMailMerge_ActivatePipe", PipeDirection.In);
                    await server.WaitForConnectionAsync(token);
                    // 收到连接请求 → 激活主窗口
                    Dispatcher.Invoke(() => ActivateMainWindow());
                    server.Disconnect();
                }
                catch (OperationCanceledException) { break; }
                catch { /* 忽略管道错误，继续监听 */ }
            }
        });
    }

    private void ActivateMainWindow()
    {
        try
        {
            if (IsIconic(new WindowInteropHelper(this).Handle))
                ShowWindow(new WindowInteropHelper(this).Handle, SW_RESTORE);
            Show();
            Activate();
            SetForegroundWindow(new WindowInteropHelper(this).Handle);
        }
        catch { }
    }

    /// <summary>
    /// 填充预览 DataGrid 的列（在试生成完成后重建）
    /// </summary>
    public void RebuildFilledPreviewColumns()
    {
        FilledPreviewDataGrid.Columns.Clear();
        var cols = _vm.FilledContentColumns;
        if (cols == null || cols.Count == 0) return;

        foreach (var colName in cols)
        {
            var column = new DataGridTextColumn
            {
                Header = colName,
                Binding = new Binding($"[{colName}]") { Mode = BindingMode.OneWay },
                MinWidth = 50
            };
            FilledPreviewDataGrid.Columns.Add(column);
        }
    }

    // 拖放：数据源区
    private void DataSource_Drop(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;
        var files = (string[])e.Data.GetData(DataFormats.FileDrop)!;
        if (files.Length > 0) _vm.LoadDataSourceFromDrop(files[0]);
    }

    // 拖放：模板区
    private void Template_Drop(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;
        var files = (string[])e.Data.GetData(DataFormats.FileDrop)!;
        if (files.Length > 0) _vm.LoadTemplateFromDrop(files[0]);
    }

    // 标题行号：仅允许输入数字
    private void NumberOnly_PreviewTextInput(object sender, TextCompositionEventArgs e)
    {
        e.Handled = !int.TryParse(e.Text, out _);
    }

    // 底部"关闭"按钮
    private void CloseApp_Click(object sender, RoutedEventArgs e)
    {
        _vm.SaveSettings();
        Close();
    }

    /// <summary>
    /// 根据 ViewModel.Columns 动态生成数据源预览 DataGrid 列
    /// </summary>
    private void RebuildPreviewColumns()
    {
        PreviewDataGrid.Columns.Clear();
        foreach (var col in _vm.Columns)
        {
            var column = new DataGridTextColumn
            {
                Header = col.DisplayName,
                Binding = new Binding($"[{col.DisplayName}]") { Mode = BindingMode.OneWay },
                MinWidth = 60
            };
            PreviewDataGrid.Columns.Add(column);
        }
    }

    /// <summary>
    /// 根据 ViewModel.TemplateContentColumns 动态生成模板内容 DataGrid 列
    /// 含占位符的单元格用黄色高亮
    /// </summary>
    private static readonly Regex PlaceholderRegex = new(@"\{(.+?)\}", RegexOptions.Compiled);

    private void RebuildTemplateContentColumns()
    {
        TemplateContentDataGrid.Columns.Clear();
        foreach (var colName in _vm.TemplateContentColumns)
        {
            var binding = new Binding($"[{colName}]") { Mode = BindingMode.OneWay };

            var column = new DataGridTextColumn
            {
                Header = colName,
                Binding = binding,
                MinWidth = 50,
                ElementStyle = CreateTemplateCellStyle()
            };
            TemplateContentDataGrid.Columns.Add(column);
        }
    }

    /// <summary>
    /// 创建模板内容单元格样式：包含占位符的单元格黄色高亮
    /// </summary>
    private Style CreateTemplateCellStyle()
    {
        // DataGridTextColumn.ElementStyle 目标类型是 TextBlock（显示模式），不是 TextBox
        var style = new Style(typeof(TextBlock));
        var trigger = new DataTrigger
        {
            Binding = new Binding(".")
            {
                Converter = new PlaceholderHighlightConverter()
            },
            Value = true
        };
        trigger.Setters.Add(new Setter(TextBlock.BackgroundProperty,
            new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 254, 243, 199)))); // #FEF3C7
        style.Triggers.Add(trigger);
        return style;
    }
}

/// <summary>
/// 值转换器：检测单元格文本是否包含 {占位符}，是则返回 true 用于高亮
/// </summary>
public class PlaceholderHighlightConverter : IValueConverter
{
    private static readonly Regex PlaceholderRegex = new(@"\{(.+?)\}", RegexOptions.Compiled);

    public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
    {
        if (value is string s && !string.IsNullOrEmpty(s))
            return PlaceholderRegex.IsMatch(s);
        return false;
    }

    public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        => throw new NotSupportedException();
}
