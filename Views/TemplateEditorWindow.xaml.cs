using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using ExcelMailMerge.Helpers;
using ExcelMailMerge.Models;
using ExcelMailMerge.Services;

namespace ExcelMailMerge.Views;

public partial class TemplateEditorWindow : Window
{
    private readonly string _templatePath;
    private readonly string _previewRoot; // 模板路径（显示 + 扫描）
    private readonly List<ColumnInfo> _columns;
    private FileSystemWatcher? _watcher;
    private DateTime _lastChange = DateTime.MinValue;
    private readonly Regex _placeholderRegex = new(@"\{.+?\}", RegexOptions.Compiled);

    public ObservableCollection<Dictionary<string, string?>> PreviewRows { get; } = new();
    public ObservableCollection<string> PreviewColumns { get; } = new();

    public string TemplatePath => _templatePath; // XAML Binding 用
    public ScannedTemplate? Scanned { get; private set; }
    public bool HasChanges { get; private set; }
    public string? SavedTemplatePath { get; private set; }

    private readonly TemplateService _tplSvc = new();

    public TemplateEditorWindow(string templatePath, List<ColumnInfo> columns)
    {
        InitializeComponent();
        DataContext = this;
        _templatePath = templatePath;
        _previewRoot = templatePath;
        _columns = columns;
        ColumnListBox.ItemsSource = columns;
        try
        {
            TplPathText.Text = templatePath;
        }
        catch { }
        LoadPreview();
        StartWatcher();
        Closed += (_, _) => StopWatcher();
    }

    // ========== Watcher ==========
    private void StartWatcher()
    {
        try
        {
            var dir = Path.GetDirectoryName(_previewRoot)!;
            var name = Path.GetFileName(_previewRoot);
            _watcher = new FileSystemWatcher(dir, name)
            {
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.CreationTime,
                EnableRaisingEvents = true
            };
            _watcher.Changed += (_, _) => QueueReload();
            _watcher.Created += (_, _) => QueueReload();
        }
        catch { }
    }
    private void StopWatcher()
    {
        if (_watcher != null)
        {
            _watcher.EnableRaisingEvents = false;
            _watcher.Dispose();
            _watcher = null;
        }
    }
    private void QueueReload()
    {
        if ((DateTime.Now - _lastChange).TotalSeconds < 1.5) return;
        _lastChange = DateTime.Now;
        _ = Task.Run(async () =>
        {
            await Task.Delay(1200);
            for (int i = 0; i < 5; i++)
            {
                try
                {
                    if (File.Exists(_previewRoot))
                    {
                        using var fs = File.Open(_previewRoot, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                        fs.Close();
                        break;
                    }
                }
                catch { await Task.Delay(400); }
            }
            System.Windows.Application.Current?.Dispatcher.BeginInvoke(() => LoadPreview());
        });
    }

    // ========== 预览 ==========
    private void LoadPreview()
    {
        PreviewRows.Clear();
        PreviewColumns.Clear();
        try
        {
            Scanned = _tplSvc.Scan(_previewRoot, null);
            var (rows, cols) = _tplSvc.GetTemplateContent(_previewRoot);
            foreach (var c in cols) PreviewColumns.Add(c);

            PreviewGrid.Columns.Clear();
            foreach (var col in PreviewColumns)
            {
                var binding = new Binding($"[{col}]") { Mode = BindingMode.OneWay };
                var column = new DataGridTextColumn
                {
                    Header = col,
                    Binding = binding,
                    MinWidth = 60,
                    ElementStyle = BuildPreviewCellStyle()
                };
                PreviewGrid.Columns.Add(column);
            }

            foreach (var row in rows)
            {
                var dict = new Dictionary<string, string?>();
                foreach (var c in PreviewColumns)
                    dict[c] = row.Values.TryGetValue(c, out var v) ? v : null;
                PreviewRows.Add(dict);
            }

            EmptyHint.Visibility = PreviewRows.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
            var phCount = Scanned?.Placeholders.Count ?? 0;
            StatText.Text = $"Sheet {Scanned?.SheetCount ?? 1} 个 · 占位符 {phCount} 个";
            HasChanges = true;
            StatusText.Text = $"✅ 预览已加载（{PreviewRows.Count} 行 × {PreviewColumns.Count} 列）";
        }
        catch (Exception ex)
        {
            StatusText.Text = $"❌ 加载失败：{ex.Message}";
            EmptyHint.Visibility = Visibility.Visible;
        }
    }

    private Style BuildPreviewCellStyle()
    {
        var style = new Style(typeof(TextBlock));
        var trigger = new DataTrigger
        {
            Binding = new Binding(".") { Converter = new PlaceholderHighlightConverter() },
            Value = true
        };
        trigger.Setters.Add(new Setter(TextBlock.BackgroundProperty,
            new SolidColorBrush(Color.FromArgb(255, 254, 243, 199)))); // #FEF3C7
        style.Triggers.Add(trigger);
        return style;
    }

    // ========== 点击列名 = 复制占位符 ==========
    private void ColumnListBox_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (ColumnListBox.SelectedItem is ColumnInfo ci) CopyPlaceholder(ci.DisplayName);
    }

    private void CopyPlaceholder_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.CommandParameter is string s) CopyPlaceholder(s);
    }
    private static void CopyPlaceholder(string colName)
    {
        try { Clipboard.SetText($"{{{colName}}}"); }
        catch { /* 静默 */ }
    }

    // ========== 搜索 ==========
    private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        var kw = SearchBox.Text?.Trim() ?? string.Empty;
        var filtered = string.IsNullOrEmpty(kw)
            ? _columns.ToList()
            : _columns.Where(c => c.DisplayName.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0).ToList();
        ColumnListBox.ItemsSource = filtered;
    }

    // ========== 按钮 ==========
    private void OpenInExcel_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            Process.Start(new ProcessStartInfo { FileName = _previewRoot, UseShellExecute = true });
            StatusText.Text = "📄 已在 Excel/WPS 中打开模板。编辑完 Ctrl+S 保存，本窗口会自动重载。";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"无法打开文件：{ex.Message}\n请确认已安装 Office 或 WPS",
                "打开失败", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }
    private void Reload_Click(object sender, RoutedEventArgs e) => LoadPreview();
    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false; Close();
    }
    private void Finish_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true; Close();
    }
}
