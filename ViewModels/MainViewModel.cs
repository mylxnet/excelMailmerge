using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Input;
using ExcelMailMerge.Helpers;
using ExcelMailMerge.Models;
using ExcelMailMerge.Services;
using Microsoft.Win32;

namespace ExcelMailMerge.ViewModels;

public class MainViewModel : ObservableObject
{
    private readonly SettingsService _settingsSvc;
    private readonly DataSourceService _dataSvc;
    private readonly TemplateService _tplSvc;
    private readonly ValidationService _valSvc;
    private readonly GenerationEngine _engine;
    private CancellationTokenSource? _cts;
    private string? _pendingPrimaryNamingCol;
    private string? _pendingSecondaryNamingCol;
    private FileSystemWatcher? _tplWatcher;
    private DateTime _lastTplChange = DateTime.MinValue;

    public MainViewModel()
    {
        _settingsSvc = ServiceLocator.Settings;
        _dataSvc = ServiceLocator.DataSource;
        _tplSvc = ServiceLocator.Template;
        _valSvc = ServiceLocator.Validation;
        _engine = ServiceLocator.Engine;

        // 初始化步骤导航
        InitStepItems();

        // 命令绑定
        BrowseDataSourceCmd = new RelayCommand(_ => BrowseDataSource());
        BrowseTemplateCmd = new RelayCommand(_ => BrowseTemplate());
        OpenTemplateEditorCmd = new RelayCommand(_ => OpenTemplateEditor(), _ => File.Exists(TemplateFilePath));
        RefreshDataSourceSheetsCmd = new RelayCommand(_ => LoadDataSourceSheets());
        ScanTemplateCmd = new RelayCommand(_ => ScanTemplate());
        ValidateCmd = new RelayCommand(_ => RunValidation(true));
        GenerateCmd = new AsyncRelayCommand(GenerateAsync, _ => CanStartGenerate);
        CancelCmd = new RelayCommand(_ => _cts?.Cancel(), _ => _cts != null && !_cts.IsCancellationRequested);
        CopyPlaceholderCmd = new RelayCommand<string>(s =>
        {
            if (!string.IsNullOrEmpty(s)) Clipboard.SetText($"{{{s}}}");
        });
        GoNextCmd = new RelayCommand(_ => GoNextStep(), _ => CanGoNext);
        GoPrevCmd = new RelayCommand(_ => GoPrevStep(), _ => CurrentStep > WizardStep.DataSource);
        GotoStepCmd = new RelayCommand<WizardStep?>(s => { if (s.HasValue) GotoStep(s.Value); });
        TrialGenerateCmd = new RelayCommand(_ => TrialGenerate(), _ => CanTrialGenerate);
        ReTrialGenerateCmd = new RelayCommand(_ => TrialGenerate(), _ => PreviewResult != null);
        ConfirmPreviewCmd = new RelayCommand(_ => GotoStep(WizardStep.Generate), _ => PreviewResult?.IsSuccess == true);
        OpenPreviewFileCmd = new RelayCommand(_ => OpenPreviewFile(), _ => PreviewResult?.IsSuccess == true);
        OpenOutputFolderCmd = new RelayCommand(_ => OpenFolder(Progress?.OutputFolder), _ => !string.IsNullOrEmpty(Progress?.OutputFolder) && Directory.Exists(Progress.OutputFolder));
        BrowseFolderCmd = new RelayCommand(_ => BrowseFolder());
        OpenTemplateInExcelCmd = new AsyncRelayCommand(OpenTemplateInExcelAsync, _ => HasTemplate);
        ReloadTemplateCmd = new RelayCommand(_ => ReloadTemplate(), _ => HasTemplate);
        FileNamePreviewMouseDownCmd = new RelayCommand(_ => { }, _ => false); // 占位，预留交互

        // 订阅 WriteSettings 变化刷新文件名预览
        WriteSettings = new AdvancedWriteSettings();
        WriteSettings.PropertyChanged += (_, _) => RefreshFileNamePreview();

        // 加载上次配置
        LoadSettings();
    }

    // ================ 向导步骤 ================
    public ObservableCollection<WizardStepItem> StepItems { get; } = new();

    private WizardStep _currentStep = WizardStep.DataSource;
    public WizardStep CurrentStep
    {
        get => _currentStep;
        set
        {
            if (SetProperty(ref _currentStep, value))
            {
                RefreshStepStatus();
                OnPropertyChanged(nameof(IsDataSourceStep));
                OnPropertyChanged(nameof(IsTemplateStep));
                OnPropertyChanged(nameof(IsPreviewStep));
                OnPropertyChanged(nameof(IsGenerateStep));
                OnPropertyChanged(nameof(CurrentStepText));
                OnPropertyChanged(nameof(CanGoNext));
                OnPropertyChanged(nameof(StageStatusText));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool IsDataSourceStep => CurrentStep == WizardStep.DataSource;
    public bool IsTemplateStep => CurrentStep == WizardStep.Template;
    public bool IsPreviewStep => CurrentStep == WizardStep.Preview;
    public bool IsGenerateStep => CurrentStep == WizardStep.Generate;

    /// <summary>当前步骤的中文显示名（用于底部状态栏）</summary>
    public string CurrentStepText => CurrentStep switch
    {
        WizardStep.DataSource => "数据源",
        WizardStep.Template => "模板",
        WizardStep.Preview => "预览",
        WizardStep.Generate => "生成",
        _ => CurrentStep.ToString()
    };

    private void InitStepItems()
    {
        StepItems.Add(new WizardStepItem { Step = WizardStep.DataSource, IndexText = "1", Label = "数据源", Desc = "上传并校验" });
        StepItems.Add(new WizardStepItem { Step = WizardStep.Template, IndexText = "2", Label = "模板", Desc = "编辑占位符" });
        StepItems.Add(new WizardStepItem { Step = WizardStep.Preview, IndexText = "3", Label = "预览", Desc = "试生成确认" });
        StepItems.Add(new WizardStepItem { Step = WizardStep.Generate, IndexText = "4", Label = "生成", Desc = "全部生成" });
        RefreshStepStatus();
    }

    private void RefreshStepStatus()
    {
        foreach (var item in StepItems)
        {
            if (item.Step < CurrentStep)
            {
                item.Status = StepStatus.Done;
                item.CanNavigate = true;
            }
            else if (item.Step == CurrentStep)
            {
                item.Status = StepStatus.Current;
                item.CanNavigate = true;
            }
            else
            {
                item.Status = StepStatus.Pending;
                item.CanNavigate = IsStepUnlocked(item.Step);
            }
            // 触发属性变化
            OnPropertyChanged(nameof(StepItems));
        }
    }

    /// <summary>
    /// 判断某步骤是否已解锁（前置条件满足）
    /// </summary>
    private bool IsStepUnlocked(WizardStep step)
    {
        return step switch
        {
            WizardStep.DataSource => true,
            WizardStep.Template => IsDataSourceValid,
            WizardStep.Preview => IsDataSourceValid && HasTemplate,
            WizardStep.Generate => IsDataSourceValid && HasTemplate && PreviewResult?.IsSuccess == true,
            _ => false
        };
    }

    /// <summary>数据源阶段是否通过（已解析且校验无错误）</summary>
    public bool IsDataSourceValid => ParsedDataSource != null && _dataSourceIssues != null && _dataSourceIssues.ErrorCount == 0;

    private ValidationResult? _dataSourceIssues;
    /// <summary>数据源阶段的校验结果（合并单元格等）</summary>
    public ValidationResult? DataSourceIssues
    {
        get => _dataSourceIssues;
        private set
        {
            SetProperty(ref _dataSourceIssues, value);
            OnPropertyChanged(nameof(IsDataSourceValid));
            OnPropertyChanged(nameof(StageStatusText));
            RefreshStepStatus();
        }
    }

    private bool CanGoNext => CurrentStep switch
    {
        WizardStep.DataSource => IsDataSourceValid,
        WizardStep.Template => HasTemplate,
        WizardStep.Preview => PreviewResult?.IsSuccess == true,
        _ => false
    };

    private void GoNextStep()
    {
        // 从模板阶段进入预览阶段：先切换步骤，再自动触发试生成
        if (CurrentStep == WizardStep.Template && PreviewResult == null && CanTrialGenerate)
        {
            CurrentStep = WizardStep.Preview;
            TrialGenerate();
            return;
        }
        if (CurrentStep < WizardStep.Generate)
            CurrentStep = CurrentStep + 1;
    }

    private void GoPrevStep()
    {
        if (CurrentStep > WizardStep.DataSource)
            CurrentStep = CurrentStep - 1;
    }

    private void GotoStep(WizardStep target)
    {
        if (!IsStepUnlocked(target) && target > CurrentStep) return;
        // 进入预览阶段时若尚未试生成，自动触发
        if (target == WizardStep.Preview && PreviewResult == null && CanTrialGenerate)
            TrialGenerate();
        CurrentStep = target;
    }

    // ================ 属性：数据源 ================
    private string? _dataSourceFilePath;
    public string? DataSourceFilePath
    {
        get => _dataSourceFilePath;
        set
        {
            SetProperty(ref _dataSourceFilePath, value);
            OnPropertyChanged(nameof(HasDataSource));
            OnPropertyChanged(nameof(StageStatusText));
            LoadDataSourceSheets();
        }
    }
    public bool HasDataSource => !string.IsNullOrWhiteSpace(DataSourceFilePath) && File.Exists(DataSourceFilePath);

    public ObservableCollection<DataSourceSheetInfo> DataSourceSheets { get; } = new();

    private DataSourceSheetInfo? _selectedSheet;
    public DataSourceSheetInfo? SelectedSheet
    {
        get => _selectedSheet;
        set
        {
            SetProperty(ref _selectedSheet, value);
            if (value != null) ParseDataSource();
        }
    }

    private int _titleRowIndex = 1;
    public int TitleRowIndex
    {
        get => _titleRowIndex;
        set
        {
            SetProperty(ref _titleRowIndex, value);
            if (value > 0) ParseDataSource();
        }
    }

    public ObservableCollection<ColumnInfo> Columns { get; } = new();
    public ObservableCollection<DataRow> PreviewRows { get; } = new();

    private ParsedDataSource? _parsedDataSource;
    public ParsedDataSource? ParsedDataSource
    {
        get => _parsedDataSource;
        private set
        {
            SetProperty(ref _parsedDataSource, value);
            OnPropertyChanged(nameof(IsDataSourceValid));
            OnPropertyChanged(nameof(StageStatusText));
            RefreshStepStatus();
        }
    }

    // ================ 属性：模板 ================
    private string? _templateFilePath;
    public string? TemplateFilePath
    {
        get => _templateFilePath;
        set
        {
            SetProperty(ref _templateFilePath, value);
            OnPropertyChanged(nameof(HasTemplate));
            OnPropertyChanged(nameof(StageStatusText));
            ScanTemplate();
            RestartTemplateWatcher();
        }
    }
    public bool HasTemplate => !string.IsNullOrWhiteSpace(TemplateFilePath) && File.Exists(TemplateFilePath);

    public ObservableCollection<TemplatePlaceholder> TemplatePlaceholders { get; } = new();

    /// <summary>模板内容预览行（展示模板中所有单元格原始内容）</summary>
    public ObservableCollection<TemplateContentRow> TemplateContentRows { get; } = new();

    /// <summary>模板内容预览列（Excel列标 A/B/C...）</summary>
    public ObservableCollection<string> TemplateContentColumns { get; } = new();

    private ScannedTemplate? _scannedTemplate;
    public ScannedTemplate? ScannedTemplate
    {
        get => _scannedTemplate;
        set
        {
            SetProperty(ref _scannedTemplate, value);
            OnPropertyChanged(nameof(StageStatusText));
            RefreshStepStatus();
        }
    }

    // ================ 属性：命名规则（两个下拉框）================
    private ColumnInfo? _primaryNamingColumn;
    public ColumnInfo? PrimaryNamingColumn
    {
        get => _primaryNamingColumn;
        set
        {
            SetProperty(ref _primaryNamingColumn, value);
            RunValidation(false);
            RefreshFileNamePreview();
        }
    }

    private ColumnInfo? _secondaryNamingColumn;
    public ColumnInfo? SecondaryNamingColumn
    {
        get => _secondaryNamingColumn;
        set
        {
            SetProperty(ref _secondaryNamingColumn, value);
            RunValidation(false);
            RefreshFileNamePreview();
        }
    }

    /// <summary>
    /// 获取当前有效的命名列列表（供校验和生成引擎使用）
    /// </summary>
    public List<string> GetNamingColumnNames()
    {
        var list = new List<string>();
        if (PrimaryNamingColumn != null) list.Add(PrimaryNamingColumn.DisplayName);
        if (SecondaryNamingColumn != null && SecondaryNamingColumn != PrimaryNamingColumn)
            list.Add(SecondaryNamingColumn.DisplayName);
        return list;
    }

    private string _namingSeparator = "_";
    public string NamingSeparator
    {
        get => _namingSeparator;
        set
        {
            SetProperty(ref _namingSeparator, string.IsNullOrEmpty(value) ? "_" : value);
            RunValidation(false);
            RefreshFileNamePreview();
        }
    }

    /// <summary>重名处理策略：自动追加序号</summary>
    public bool AutoAppendIndex { get; set; } = true;

    /// <summary>文件名预览文本</summary>
    private string _fileNamePreview = "（请先选择命名列）";
    public string FileNamePreview
    {
        get => _fileNamePreview;
        private set
        {
            if (SetProperty(ref _fileNamePreview, value))
                OnPropertyChanged(nameof(StageStatusText));
        }
    }

    /// <summary>
    /// 底部状态栏文本（按当前阶段显示不同提示，避免在数据源/模板阶段显示不相关的命名列提示）
    /// </summary>
    public string StageStatusText => CurrentStep switch
    {
        WizardStep.DataSource => IsDataSourceValid ? "✅ 数据源已校验通过，可进入下一步" : (HasDataSource ? "⚠️ 数据源有错误，请先修复" : "请加载数据源文件"),
        WizardStep.Template => HasTemplate ? $"模板已加载：{ScannedTemplate?.Placeholders.Count ?? 0} 个占位符，未匹配 {ScannedTemplate?.UnmatchedCount ?? 0} 个" : "请加载模板文件",
        WizardStep.Preview => PreviewResult?.IsSuccess == true ? $"试生成成功：{PreviewResult.PlaceholderReplacedCount} 个占位符已替换，{PreviewResult.FormulaConvertedCount} 个公式已转值" : "请先试生成确认效果",
        WizardStep.Generate => FileNamePreview,
        _ => string.Empty
    };

    /// <summary>
    /// 刷新文件名预览
    ///   多文件模式：每行一个 .xlsx，列前 3 个文件名示例
    ///   单文件模式：输出 1 个 .xlsx，文件名固定为"合并结果_时间戳.xlsx"；右侧预览各 Sheet 名
    /// </summary>
    public void RefreshFileNamePreview()
    {
        if (ParsedDataSource == null || ParsedDataSource.Rows.Count == 0 || PrimaryNamingColumn == null)
        {
            FileNamePreview = PrimaryNamingColumn == null ? "（请先选择命名列）" : "（无数据）";
            return;
        }
        try
        {
            var names = GetNamingColumnNames();
            var sep = NamingSeparator;
            var naming = ServiceLocator.Naming;
            int show = Math.Min(3, ParsedDataSource.Rows.Count);

            if (IsSingleFileMode)
            {
                // 单文件模式：文件名 = "合并结果_yyyyMMdd_HHmmss.xlsx"（与 GenerationEngine 硬编码一致）
                string fileBase = string.IsNullOrWhiteSpace(SingleFileName)
                    ? $"合并结果_{DateTime.Now:yyyyMMdd_HHmmss}"
                    : naming.SanitizeForFileName(SingleFileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                        ? SingleFileName[..^5] : SingleFileName);
                // 预览 Sheet 名（前 3 个）
                var sheetParts = new List<string>();
                for (int i = 0; i < show; i++)
                {
                    var raw = naming.BuildRawName(ParsedDataSource.Rows[i], names, sep);
                    var safe = naming.SanitizeForFileName(raw);
                    if (!string.IsNullOrEmpty(safe)) sheetParts.Add(safe);
                }
                FileNamePreview = $"{fileBase}.xlsx　|　Sheet: {string.Join(" | ", sheetParts)}";
            }
            else
            {
                // 多文件模式：每行一个工作簿
                var sb = new System.Text.StringBuilder();
                for (int i = 0; i < show; i++)
                {
                    var raw = naming.BuildRawName(ParsedDataSource.Rows[i], names, sep);
                    var safe = naming.SanitizeForFileName(raw);
                    if (string.IsNullOrEmpty(safe)) safe = "（空命名）";
                    sb.Append(safe).Append(".xlsx");
                    if (i < show - 1) sb.Append("  |  ");
                }
                FileNamePreview = sb.ToString();
            }
        }
        catch { FileNamePreview = "（预览失败）"; }
    }

    // ================ 属性：输出模式 ================
    private OutputMode _outputMode = OutputMode.MultiFilePerRow;
    public OutputMode OutputMode
    {
        get => _outputMode;
        set
        {
            SetProperty(ref _outputMode, value);
            OnPropertyChanged(nameof(IsSingleFileMode));
            RefreshFileNamePreview();
            RunValidation(false);
        }
    }
    public bool IsSingleFileMode => OutputMode == OutputMode.SingleFileMultiSheet;

    private string? _singleFileName;
    public string? SingleFileName { get => _singleFileName; set => SetProperty(ref _singleFileName, value); }

    private string? _customOutputFolder;
    public string? CustomOutputFolder { get => _customOutputFolder; set => SetProperty(ref _customOutputFolder, value); }

    // ================ 属性：高级设置 ================
    private AdvancedWriteSettings _writeSettings = new();
    public AdvancedWriteSettings WriteSettings { get => _writeSettings; set => SetProperty(ref _writeSettings, value); }

    // ================ 属性：校验（生成阶段总校验）================
    private ValidationResult? _validationResult;
    public ValidationResult? ValidationResult
    {
        get => _validationResult;
        set
        {
            SetProperty(ref _validationResult, value);
            OnPropertyChanged(nameof(CanStartGenerate));
        }
    }

    public bool CanStartGenerate => ValidationResult != null && ValidationResult.IsSuccess;

    // ================ 属性：试生成预览 ================
    private PreviewResult? _previewResult;
    public PreviewResult? PreviewResult
    {
        get => _previewResult;
        private set
        {
            SetProperty(ref _previewResult, value);
            OnPropertyChanged(nameof(HasPreviewResult));
            OnPropertyChanged(nameof(FilledContentRows));
            OnPropertyChanged(nameof(FilledContentColumns));
            OnPropertyChanged(nameof(StageStatusText));
            RefreshStepStatus();
            CommandManager.InvalidateRequerySuggested();
        }
    }
    public bool HasPreviewResult => PreviewResult != null;

    /// <summary>填充了数据的模板内容预览行</summary>
    public List<TemplateContentRow> FilledContentRows => PreviewResult?.FilledContentRows ?? new();

    /// <summary>填充了数据的模板内容预览列</summary>
    public List<string> FilledContentColumns => PreviewResult?.FilledContentColumns ?? new();

    /// <summary>是否可执行试生成</summary>
    public bool CanTrialGenerate => ParsedDataSource != null && ParsedDataSource.Rows.Count > 0
        && ScannedTemplate != null;

    // ================ 属性：生成进度 ================
    private GenerationProgress? _progress;
    public GenerationProgress? Progress { get => _progress; set => SetProperty(ref _progress, value); }

    private bool _isGenerating;
    public bool IsGenerating
    {
        get => _isGenerating;
        set
        {
            SetProperty(ref _isGenerating, value);
            CommandManager.InvalidateRequerySuggested();
        }
    }

    // ================ 命令 ================
    public ICommand BrowseDataSourceCmd { get; }
    public ICommand BrowseTemplateCmd { get; }
    public ICommand OpenTemplateEditorCmd { get; }
    public ICommand RefreshDataSourceSheetsCmd { get; }
    public ICommand ScanTemplateCmd { get; }
    public ICommand ValidateCmd { get; }
    public ICommand GenerateCmd { get; }
    public ICommand CancelCmd { get; }
    public ICommand CopyPlaceholderCmd { get; }
    public ICommand GoNextCmd { get; }
    public ICommand GoPrevCmd { get; }
    public ICommand GotoStepCmd { get; }
    public ICommand TrialGenerateCmd { get; }
    public ICommand ReTrialGenerateCmd { get; }
    public ICommand ConfirmPreviewCmd { get; }
    public ICommand OpenPreviewFileCmd { get; }
    public ICommand OpenOutputFolderCmd { get; }
    public ICommand BrowseFolderCmd { get; }
    public ICommand FileNamePreviewMouseDownCmd { get; }

    // ================ 方法：浏览文件 ================
    private void BrowseDataSource()
    {
        var dlg = new OpenFileDialog { Filter = "Excel文件 (*.xlsx)|*.xlsx" };
        if (dlg.ShowDialog() == true) DataSourceFilePath = dlg.FileName;
    }
    private void BrowseTemplate()
    {
        var dlg = new OpenFileDialog { Filter = "Excel模板 (*.xlsx)|*.xlsx" };
        if (dlg.ShowDialog() == true) TemplateFilePath = dlg.FileName;
    }

    public event Action<string>? RequestSetTemplateFilePath;

    private void OpenTemplateEditor()
    {
        if (string.IsNullOrEmpty(TemplateFilePath)) return;
        var editor = new Views.TemplateEditorWindow(TemplateFilePath, Columns.ToList());
        if (editor.ShowDialog() == true && !string.IsNullOrEmpty(editor.SavedTemplatePath))
        {
            TemplateFilePath = editor.SavedTemplatePath;
            RequestSetTemplateFilePath?.Invoke(editor.SavedTemplatePath);
            MessageBox.Show("✅ 模板编辑完成，已自动回填为当前模板文件。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    // ================ 方法：浏览文件夹 ================
    private void BrowseFolder()
    {
        // 用 OpenFileDialog 选任意文件取所在目录（避免 Windows Forms 依赖）
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Title = "选择保存位置（在任意位置选一个文件即可，将使用其所在目录）",
            Filter = "所有文件|*.*",
            CheckFileExists = false,
            FileName = "选择此文件夹"
        };
        if (!string.IsNullOrEmpty(CustomOutputFolder) && Directory.Exists(CustomOutputFolder))
            dialog.InitialDirectory = CustomOutputFolder;
        else if (HasDataSource)
        {
            var dir = Path.GetDirectoryName(DataSourceFilePath);
            if (!string.IsNullOrEmpty(dir)) dialog.InitialDirectory = dir;
        }
        if (dialog.ShowDialog() == true)
        {
            var dir = Path.GetDirectoryName(dialog.FileName);
            if (!string.IsNullOrEmpty(dir))
                CustomOutputFolder = dir;
        }
    }

    // ================ 方法：数据源 ================
    public void LoadDataSourceFromDrop(string path)
    {
        if (File.Exists(path) && (path.EndsWith(".xlsx") || path.EndsWith(".xlsm")))
            DataSourceFilePath = path;
    }

    private void LoadDataSourceSheets()
    {
        DataSourceSheets.Clear();
        if (!HasDataSource) return;
        try
        {
            var sheets = _dataSvc.GetSheetInfos(DataSourceFilePath!);
            foreach (var s in sheets) DataSourceSheets.Add(s);
            var first = sheets.FirstOrDefault(s => s.HasData) ?? sheets.FirstOrDefault();
            if (first != null)
            {
                try { TitleRowIndex = _dataSvc.AutoDetectTitleRow(DataSourceFilePath!, first.SheetName); }
                catch { TitleRowIndex = 1; }
                _selectedSheet = first;
                OnPropertyChanged(nameof(SelectedSheet));
                ParseDataSource();
            }
        }
        catch (Exception ex) { MessageBox.Show($"读取数据源Sheet失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error); }
    }

    private void ParseDataSource()
    {
        PreviewRows.Clear();
        Columns.Clear();
        ParsedDataSource = null;
        DataSourceIssues = null;
        if (!HasDataSource || SelectedSheet == null) return;
        try
        {
            var parsed = _dataSvc.Parse(DataSourceFilePath!, SelectedSheet.SheetName, TitleRowIndex <= 0 ? 1 : TitleRowIndex);
            ParsedDataSource = parsed;
            foreach (var c in parsed.Columns) Columns.Add(c);
            foreach (var r in parsed.Rows.Take(5)) PreviewRows.Add(r);

            // 数据源阶段自动校验（合并单元格、标题等）
            DataSourceIssues = _valSvc.ValidateDataSource(parsed);

            // 恢复上次选中的命名列
            if (_pendingPrimaryNamingCol != null)
            {
                PrimaryNamingColumn = parsed.Columns.FirstOrDefault(c => c.DisplayName == _pendingPrimaryNamingCol);
                _pendingPrimaryNamingCol = null;
            }
            if (_pendingSecondaryNamingCol != null)
            {
                SecondaryNamingColumn = parsed.Columns.FirstOrDefault(c => c.DisplayName == _pendingSecondaryNamingCol);
                _pendingSecondaryNamingCol = null;
            }
            if (HasTemplate) ScanTemplate();
            RefreshFileNamePreview();
            RefreshStepStatus();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"解析数据源失败：{ex.Message}\n\n请确认：\n1. 标题行号是否正确\n2. 数据区域是否有合并单元格\n3. 列名中是否有大括号{{}}",
                "解析失败", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // ================ 方法：模板扫描 ================
    public void LoadTemplateFromDrop(string path)
    {
        if (File.Exists(path) && (path.EndsWith(".xlsx") || path.EndsWith(".xlsm")))
            TemplateFilePath = path;
    }

    private void ScanTemplate()
    {
        TemplatePlaceholders.Clear();
        TemplateContentRows.Clear();
        TemplateContentColumns.Clear();
        ScannedTemplate = null;
        if (!HasTemplate) return;
        try
        {
            var scanned = _tplSvc.Scan(TemplateFilePath!, ParsedDataSource);
            ScannedTemplate = scanned;
            foreach (var p in scanned.Placeholders) TemplatePlaceholders.Add(p);
            foreach (var c in Columns) OnPropertyChanged(nameof(c.IsUsedInTemplate));

            // 读取模板内容预览
            var (contentRows, contentColumns) = _tplSvc.GetTemplateContent(TemplateFilePath!);
            foreach (var col in contentColumns) TemplateContentColumns.Add(col);
            foreach (var row in contentRows) TemplateContentRows.Add(row);

            // 进入模板阶段时清除旧的预览结果
            if (PreviewResult != null)
            {
                _engine.ClearPreviewFile(PreviewResult.TempFilePath);
                PreviewResult = null;
            }
            RunValidation(false);
            RefreshStepStatus();
        }
        catch (Exception ex) { MessageBox.Show($"扫描模板失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error); }
    }

    // ================ 方法：模板 Watcher（文件变更自动重载） ================
    private void RestartTemplateWatcher()
    {
        StopTemplateWatcher();
        if (!HasTemplate) return;
        try
        {
            var dir = Path.GetDirectoryName(TemplateFilePath!)!;
            var name = Path.GetFileName(TemplateFilePath!);
            _tplWatcher = new FileSystemWatcher(dir, name)
            {
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.CreationTime | NotifyFilters.LastAccess,
                EnableRaisingEvents = true
            };
            _tplWatcher.Changed += (_, _) => QueueReloadTemplate();
            _tplWatcher.Created += (_, _) => QueueReloadTemplate();
            _tplWatcher.Renamed += (_, _) => QueueReloadTemplate();
        }
        catch { /* Watcher 启动失败不影响主流程 */ }
    }
    private void StopTemplateWatcher()
    {
        if (_tplWatcher != null)
        {
            _tplWatcher.EnableRaisingEvents = false;
            _tplWatcher.Dispose();
            _tplWatcher = null;
        }
    }
    private void QueueReloadTemplate()
    {
        var now = DateTime.Now;
        if ((now - _lastTplChange).TotalSeconds < 1.5) return; // debounce
        _lastTplChange = now;
        // 文件刚保存时可能还被 Excel 锁着，延迟尝试 + 重试
        _ = Task.Run(async () =>
        {
            await Task.Delay(1200);
            for (int retry = 0; retry < 5; retry++)
            {
                try
                {
                    if (TemplateFilePath != null && File.Exists(TemplateFilePath))
                    {
                        using var fs = File.Open(TemplateFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                        fs.Close();
                        break; // 文件锁释放
                    }
                }
                catch { await Task.Delay(400); }
            }
            System.Windows.Application.Current?.Dispatcher.BeginInvoke(() =>
            {
                if (HasTemplate) ScanTemplate();
            });
        });
    }

    // ================ 命令：在 Excel/WPS 中打开模板 ================
    public AsyncRelayCommand? OpenTemplateInExcelCmd { get; private set; }
    private async Task OpenTemplateInExcelAsync(object? _)
    {
        if (!HasTemplate) return;
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = TemplateFilePath!,
                UseShellExecute = true
            });
            await Task.Delay(800);
            RestartTemplateWatcher();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"无法打开文件：{ex.Message}\n请确认已安装 Office 或 WPS",
                "打开失败", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    // ================ 命令：手动重载模板预览 ================
    public RelayCommand? ReloadTemplateCmd { get; private set; }
    private void ReloadTemplate()
    {
        if (!HasTemplate) return;
        ScanTemplate();
    }

    // ================ 方法：校验 ================
    private void RunValidation(bool showMsgOnSuccess)
    {
        if (ParsedDataSource == null || ScannedTemplate == null) { ValidationResult = null; return; }
        try
        {
            ValidationResult = _valSvc.Validate(ParsedDataSource, ScannedTemplate,
                GetNamingColumnNames(), NamingSeparator, OutputMode);
            if (showMsgOnSuccess)
                MessageBox.Show(ValidationResult.SummaryText,
                    ValidationResult.IsSuccess ? "✅ 校验通过" : "❌ 校验未通过",
                    MessageBoxButton.OK,
                    ValidationResult.IsSuccess ? MessageBoxImage.Information : MessageBoxImage.Warning);
        }
        catch (Exception ex) { MessageBox.Show($"校验异常：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error); }
    }

    // ================ 方法：试生成预览 ================
    private void TrialGenerate()
    {
        if (!CanTrialGenerate) return;
        try
        {
            // 清除旧预览
            if (PreviewResult != null)
                _engine.ClearPreviewFile(PreviewResult.TempFilePath);

            var result = _engine.GeneratePreview(ParsedDataSource!, ScannedTemplate!, WriteSettings);
            PreviewResult = result;
            if (!result.IsSuccess)
                MessageBox.Show($"试生成失败：{result.ErrorMessage}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"试生成异常：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OpenPreviewFile()
    {
        if (PreviewResult?.IsSuccess != true || !File.Exists(PreviewResult.TempFilePath)) return;
        try { Process.Start("explorer.exe", PreviewResult.TempFilePath); }
        catch { /* 忽略 */ }
    }

    private void OpenFolder(string? path)
    {
        if (string.IsNullOrEmpty(path) || !Directory.Exists(path)) return;
        try { Process.Start("explorer.exe", path); }
        catch { /* 忽略 */ }
    }

    // ================ 方法：生成 ================
    private async Task GenerateAsync(object? _)
    {
        if (!CanStartGenerate || ParsedDataSource == null || ScannedTemplate == null) return;
        IsGenerating = true;
        _cts = new CancellationTokenSource();
        Progress = new GenerationProgress { TotalRows = ParsedDataSource.Rows.Count };
        var pg = Progress;
        var progressImpl = new Progress<GenerationProgress>(p =>
        {
            pg.CurrentRowIndex = p.CurrentRowIndex;
            pg.SuccessCount = p.SuccessCount;
            pg.FailCount = p.FailCount;
            pg.SkipCount = p.SkipCount;
            pg.IsCompleted = p.IsCompleted;
            pg.IsCanceled = p.IsCanceled;
            pg.OutputFolder = p.OutputFolder;
            if (p.Logs.Count > pg.Logs.Count)
            {
                for (int i = pg.Logs.Count; i < p.Logs.Count; i++) pg.Logs.Add(p.Logs[i]);
            }
            OnPropertyChanged(nameof(Progress));
        });

        try
        {
            // 清除临时预览文件
            if (PreviewResult != null)
                _engine.ClearPreviewFile(PreviewResult.TempFilePath);

            var result = await _engine.GenerateAsync(
                ParsedDataSource, ScannedTemplate,
                GetNamingColumnNames(), NamingSeparator,
                OutputMode, CustomOutputFolder, WriteSettings,
                SingleFileName, progressImpl, _cts.Token);

            // 生成完成后自动打开输出目录（与需求"生成后自动打开输出目录"对应）
            if (!result.IsCanceled && result.FailCount == 0)
                OpenFolder(result.OutputFolder);

            MessageBox.Show(result.StatusText + $"\n\n输出目录：{result.OutputFolder}",
                result.IsCanceled ? "已取消" : (result.FailCount == 0 ? "✅ 完成" : "⚠️ 有失败"),
                MessageBoxButton.OK,
                result.IsCanceled ? MessageBoxImage.Warning : (result.FailCount == 0 ? MessageBoxImage.Information : MessageBoxImage.Warning));
            SaveSettings();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"生成失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            _cts?.Dispose();
            _cts = null;
            IsGenerating = false;
        }
    }

    // ================ 配置持久化 ================
    private void LoadSettings()
    {
        var s = _settingsSvc.Load();
        if (s == null) return;
        _dataSourceFilePath = s.DataSourceFilePath; OnPropertyChanged(nameof(DataSourceFilePath));
        _titleRowIndex = s.TitleRowIndex < 1 ? 1 : s.TitleRowIndex; OnPropertyChanged(nameof(TitleRowIndex));
        _templateFilePath = s.TemplateFilePath; OnPropertyChanged(nameof(TemplateFilePath));
        _outputMode = s.OutputMode; OnPropertyChanged(nameof(OutputMode)); OnPropertyChanged(nameof(IsSingleFileMode));
        _namingSeparator = string.IsNullOrEmpty(s.NamingSeparator) ? "_" : s.NamingSeparator; OnPropertyChanged(nameof(NamingSeparator));
        _customOutputFolder = s.CustomOutputFolder; OnPropertyChanged(nameof(CustomOutputFolder));
        _writeSettings = s.WriteSettings ?? new AdvancedWriteSettings();
        _writeSettings.PropertyChanged += (_, _) => RefreshFileNamePreview();
        OnPropertyChanged(nameof(WriteSettings));
        _pendingPrimaryNamingCol = s.PrimaryNamingColumn;
        _pendingSecondaryNamingCol = s.SecondaryNamingColumn;

        Application.Current?.Dispatcher.BeginInvoke(() =>
        {
            if (!string.IsNullOrWhiteSpace(DataSourceFilePath) && File.Exists(DataSourceFilePath))
                LoadDataSourceSheets();
            else if (!string.IsNullOrWhiteSpace(TemplateFilePath) && File.Exists(TemplateFilePath))
                ScanTemplate();
        });
    }

    public void SaveSettings()
    {
        _settingsSvc.Save(new AppSettings
        {
            DataSourceFilePath = DataSourceFilePath,
            DataSourceSheetName = SelectedSheet?.SheetName,
            TitleRowIndex = TitleRowIndex,
            TemplateFilePath = TemplateFilePath,
            OutputMode = OutputMode,
            PrimaryNamingColumn = PrimaryNamingColumn?.DisplayName,
            SecondaryNamingColumn = SecondaryNamingColumn?.DisplayName,
            NamingSeparator = NamingSeparator,
            CustomOutputFolder = CustomOutputFolder,
            WriteSettings = WriteSettings
        });
    }
}
