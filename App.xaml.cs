using System.IO;
using System.Text;
using System.Windows;
using ExcelMailMerge.Models;
using ExcelMailMerge.Services;

namespace ExcelMailMerge;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        try
        {
            // 命令行模式：--gen-test-files 生成测试数据文件后退出
            if (e.Args.Length > 0 && e.Args[0] == "--gen-test-files")
            {
                GenTestFiles();
                Shutdown(0);
                return;
            }

            // 命令行模式：--load-test 测试加载已有 xlsx 文件是否崩溃
            if (e.Args.Length > 0 && e.Args[0] == "--load-test")
            {
                LoadTest();
                Shutdown(0);
                return;
            }

            // 命令行模式：--smoke-test 端到端业务逻辑测试
            if (e.Args.Length > 0 && e.Args[0] == "--smoke-test")
            {
                SmokeTest();
                Shutdown(0);
                return;
            }

            // 正常启动：显示主窗口
            base.OnStartup(e);
            var mainWin = new MainWindow();
            mainWin.Show();
        }
        catch (Exception ex)
        {
            var logPath = Path.Combine(Path.GetTempPath(), "ExcelMailMerge_startup_error.txt");
            File.WriteAllText(logPath, FlattenException(ex));
            MessageBox.Show($"{ex.GetType().Name}: {ex.Message}\n\n详情已写入: {logPath}", "启动失败", MessageBoxButton.OK, MessageBoxImage.Error);
            Shutdown(1);
        }
    }

    /// <summary>
    /// 生成测试用数据源和模板文件（命令行 --gen-test-files 触发）
    /// </summary>
    private static void GenTestFiles()
    {
        var outDir = Path.Combine(AppContext.BaseDirectory, "测试数据");
        Directory.CreateDirectory(outDir);
        var logFile = Path.Combine(Path.GetTempPath(), "ExcelMailMerge_gentest.log");
        var log = new StringBuilder();
        void Log(string s) { log.AppendLine(s); }
        Log($"输出目录: {outDir}");

        try
        {
            // 1. 数据源：员工信息表
            var dataSourcePath = Path.Combine(outDir, "数据源_员工信息.xlsx");
            using (var wb = new ClosedXML.Excel.XLWorkbook())
            {
                var ws = wb.AddWorksheet("员工信息");
                string[] headers = { "工号", "姓名", "部门", "职位", "入职日期", "手机号", "基本工资", "绩效系数" };
                for (int c = 0; c < headers.Length; c++)
                    ws.Cell(1, c + 1).Value = headers[c];

                object?[][] data =
                {
                    new object?[] { "A001", "张三", "技术部", "高级工程师", new DateTime(2020, 3, 15), "13800138000", 15000.00m, 1.2 },
                    new object?[] { "A002", "李四", "市场部", "市场经理", new DateTime(2019, 7, 1), "13900139000", 18000.00m, 1.5 },
                    new object?[] { "A003", "王五", "财务部", "会计主管", new DateTime(2021, 1, 10), "13700137000", 13000.00m, 1.1 },
                    new object?[] { "A004", "赵六", "人事部", "HR专员", new DateTime(2022, 5, 20), "13600136000", 10000.00m, 1.0 },
                };
                for (int r = 0; r < data.Length; r++)
                    for (int c = 0; c < data[r].Length; c++)
                    {
                        var cell = ws.Cell(r + 2, c + 1);
                        var v = data[r][c];
                        if (v is DateTime dt) cell.Value = dt;
                        else if (v is decimal dec) cell.Value = dec;
                        else if (v is double d) cell.Value = d;
                        else cell.Value = v?.ToString() ?? "";
                    }
                ws.Row(1).Style.Font.Bold = true;
                ws.Row(1).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#D4A256");
                ws.Row(1).Style.Font.FontColor = ClosedXML.Excel.XLColor.White;
                // 固定列宽（避免 AdjustToContents 因 SixLabors.Fonts 版本不兼容而崩溃）
                for (int c = 1; c <= headers.Length; c++)
                    ws.Column(c).Width = 16;
                wb.SaveAs(dataSourcePath);
            }
            Log($"数据源生成: {dataSourcePath}");

            // 2. 模板：员工信息卡（含占位符 + 公式）
            var templatePath = Path.Combine(outDir, "模板_员工信息卡.xlsx");
            using (var wb = new ClosedXML.Excel.XLWorkbook())
            {
                var ws = wb.AddWorksheet("信息卡");
                ws.Cell(1, 1).Value = "员工信息卡";
                ws.Range(1, 1, 1, 4).Merge();
                ws.Cell(1, 1).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                ws.Cell(1, 1).Style.Font.Bold = true;
                ws.Cell(1, 1).Style.Font.FontSize = 16;
                ws.Cell(1, 1).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#1A2540");
                ws.Cell(1, 1).Style.Font.FontColor = ClosedXML.Excel.XLColor.White;

                ws.Cell(2, 1).Value = "工号：";   ws.Cell(2, 2).Value = "{工号}";
                ws.Cell(2, 3).Value = "姓名：";   ws.Cell(2, 4).Value = "{姓名}";
                ws.Cell(3, 1).Value = "部门：";   ws.Cell(3, 2).Value = "{部门}";
                ws.Cell(3, 3).Value = "职位：";   ws.Cell(3, 4).Value = "{职位}";
                ws.Cell(4, 1).Value = "入职日期："; ws.Cell(4, 2).Value = "{入职日期}";
                ws.Cell(4, 3).Value = "手机号：";  ws.Cell(4, 4).Value = "{手机号}";
                ws.Cell(5, 1).Value = "基本工资："; ws.Cell(5, 2).Value = "{基本工资}";
                ws.Cell(5, 3).Value = "绩效系数："; ws.Cell(5, 4).Value = "{绩效系数}";
                ws.Cell(6, 1).Value = "实发工资：";
                ws.Cell(6, 2).FormulaA1 = "=B5*D5";

                ws.Column(1).Width = 14; ws.Column(2).Width = 20;
                ws.Column(3).Width = 14; ws.Column(4).Width = 20;
                for (int r = 2; r <= 6; r++)
                    for (int c = 1; c <= 4; c++)
                        ws.Cell(r, c).Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                wb.SaveAs(templatePath);
            }
            Log($"模板生成: {templatePath}");

            // 3. 错误数据源（含合并单元格）
            var badPath = Path.Combine(outDir, "数据源_错误_含合并单元格.xlsx");
            using (var wb = new ClosedXML.Excel.XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Cell(1, 1).Value = "标题A"; ws.Cell(1, 2).Value = "标题B"; ws.Cell(1, 3).Value = "标题C";
                ws.Range(2, 1, 2, 2).Merge();
                ws.Cell(2, 1).Value = "数据1"; ws.Cell(2, 3).Value = "数据3";
                ws.Cell(3, 1).Value = "数据1-2"; ws.Cell(3, 2).Value = "数据2-2"; ws.Cell(3, 3).Value = "数据3-2";
                wb.SaveAs(badPath);
            }
            Log($"错误数据源生成: {badPath}");
            Log("全部测试文件生成完成！");
        }
        catch (Exception ex)
        {
            Log($"ERROR: {ex.GetType().FullName}: {ex.Message}");
            Log(ex.StackTrace ?? "");
            Log($"Inner: {ex.InnerException?.Message}");
        }
        finally
        {
            File.WriteAllText(logFile, log.ToString());
        }
    }

    /// <summary>
    /// 测试加载已有 xlsx 文件是否因 SixLabors.Fonts 版本不兼容而崩溃
    /// </summary>
    private static void LoadTest()
    {
        var logFile = Path.Combine(Path.GetTempPath(), "ExcelMailMerge_loadtest.log");
        var log = new StringBuilder();
        var testDir = Path.Combine(AppContext.BaseDirectory, "测试数据");
        var files = new[] { "数据源_员工信息.xlsx", "模板_员工信息卡.xlsx", "数据源_错误_含合并单元格.xlsx" };

        foreach (var f in files)
        {
            var path = Path.Combine(testDir, f);
            log.AppendLine($"--- 加载: {path} ---");
            try
            {
                using var wb = new ClosedXML.Excel.XLWorkbook(path);
                log.AppendLine($"  ✓ 成功: Worksheets={wb.Worksheets.Count}");
                foreach (var ws in wb.Worksheets)
                {
                    log.AppendLine($"    Sheet: {ws.Name}, 行数={ws.LastRowUsed()?.RowNumber() ?? 0}, 列数={ws.LastColumnUsed()?.ColumnNumber() ?? 0}");
                    // 尝试读取单元格值（触发潜在的字体计算）
                    int cellCount = 0;
                    foreach (var cell in ws.CellsUsed())
                    {
                        var v = cell.Value;
                        cellCount++;
                        if (cellCount > 20) break;
                    }
                    log.AppendLine($"    读取单元格数: {cellCount}");
                }
            }
            catch (Exception ex)
            {
                log.AppendLine($"  ✗ 失败: {ex.GetType().FullName}: {ex.Message}");
                if (ex.InnerException != null)
                    log.AppendLine($"    Inner: {ex.InnerException.GetType().FullName}: {ex.InnerException.Message}");
                log.AppendLine($"    Stack: {ex.StackTrace?.Substring(0, Math.Min(500, ex.StackTrace?.Length ?? 0))}");
            }
            log.AppendLine();
        }
        File.WriteAllText(logFile, log.ToString());
    }

    /// <summary>
    /// 端到端冒烟测试：数据源解析→模板扫描→校验→试生成→全部生成
    /// </summary>
    private static void SmokeTest()
    {
        var logFile = Path.Combine(Path.GetTempPath(), "ExcelMailMerge_smoketest.log");
        var log = new StringBuilder();
        void Log(string s) { log.AppendLine(s); }
        var testDir = Path.Combine(AppContext.BaseDirectory, "测试数据");
        var dataSourcePath = Path.Combine(testDir, "数据源_员工信息.xlsx");
        var templatePath = Path.Combine(testDir, "模板_员工信息卡.xlsx");
        var outputDir = Path.Combine(testDir, "输出结果");

        try
        {
            Log("========== 冒烟测试开始 ==========");
            Log($"数据源: {dataSourcePath}");
            Log($"模板: {templatePath}");

            // 1. 数据源解析
            Log("\n--- 1. 数据源解析 ---");
            var ds = ServiceLocator.DataSource;
            var parsed = ds.Parse(dataSourcePath, "员工信息", 1);
            Log($"  列数: {parsed.Columns.Count}");
            Log($"  数据行数: {parsed.Rows.Count}");
            foreach (var col in parsed.Columns)
                Log($"  列: [{col.DisplayName}] (原名:{col.OriginalName})");
            for (int i = 0; i < Math.Min(3, parsed.Rows.Count); i++)
            {
                var r = parsed.Rows[i];
                var vals = r.Values.Select(kv => $"{kv.Key}={kv.Value}").ToList();
                Log($"  行{i+1} (原始行{r.OriginalRowIndex}): {string.Join(" | ", vals)}");
            }

            // 2. 模板扫描
            Log("\n--- 2. 模板扫描 ---");
            var ts = ServiceLocator.Template;
            var template = ts.Scan(templatePath, parsed);
            Log($"  Sheet数: {template.SheetCount}");
            Log($"  占位符总数: {template.Placeholders.Count}");
            Log($"  已匹配: {template.Placeholders.Count(p => p.IsMatched)}");
            Log($"  未匹配: {template.UnmatchedCount}");
            foreach (var sn in template.SheetNames)
                Log($"  Sheet名: {sn}");
            foreach (var p in template.Placeholders)
                Log($"  占位符: {p.FullText} 出现{p.OccurrenceCount}次 匹配={p.IsMatched} {p.MatchStatusText}");

            // 3. 校验
            Log("\n--- 3. 数据源校验 ---");
            var vs = ServiceLocator.Validation;
            var vResult = vs.ValidateDataSource(parsed);
            Log($"  问题数: {vResult.Issues.Count}");
            foreach (var issue in vResult.Issues)
                Log($"  [{issue.Level}] {issue.Message}");

            // 4. 试生成预览
            Log("\n--- 4. 试生成预览 ---");
            var engine = ServiceLocator.Engine;
            var writeSettings = new AdvancedWriteSettings();
            var preview = engine.GeneratePreview(parsed, template, writeSettings);
            Log($"  成功: {preview.IsSuccess}");
            Log($"  临时文件: {preview.TempFilePath}");
            Log($"  Sheet数: {preview.SheetCount}");
            Log($"  占位符替换: {preview.PlaceholderReplacedCount}");
            Log($"  公式转值: {preview.FormulaConvertedCount}");
            Log($"  填充预览行数: {preview.FilledContentRows.Count}，列数: {preview.FilledContentColumns.Count}");
            if (!preview.IsSuccess)
                Log($"  错误: {preview.ErrorMessage}");
            if (!string.IsNullOrEmpty(preview.TempFilePath) && File.Exists(preview.TempFilePath))
                Log($"  临时文件存在: ✓ ({new FileInfo(preview.TempFilePath).Length} bytes)");

            // 5. 全部生成（多文件模式）
            Log("\n--- 5. 全部生成（多文件模式） ---");
            // 清理旧输出
            if (Directory.Exists(outputDir))
                Directory.Delete(outputDir, true);
            Directory.CreateDirectory(outputDir);

            var namingCols = new List<string> { "工号", "姓名" };
            var progress = new Progress<GenerationProgress>(p =>
                Log($"  进度: {p.CurrentRowIndex}/{p.TotalRows} 成功={p.SuccessCount} 失败={p.FailCount}"));

            // 注意：不能用 .Result 直接阻塞主线程——GenerateAsync 内部有 await Task.Yield()
            // 会捕获 Dispatcher SynchronizationContext，导致死锁。用 Task.Run 切换到线程池（无 SyncContext）。
            var genResult = Task.Run(() => engine.GenerateAsync(
                parsed, template, namingCols, "_",
                OutputMode.MultiFilePerRow, outputDir,
                writeSettings, null, progress)).GetAwaiter().GetResult();

            Log($"  完成: {genResult.IsCompleted}");
            Log($"  成功数: {genResult.SuccessCount}");
            Log($"  失败数: {genResult.FailCount}");
            Log($"  输出目录: {genResult.OutputFolder}");
            if (genResult.Failures.Count > 0)
                foreach (var f in genResult.Failures)
                    Log($"  失败行{f.rowIndex} ({f.fileName}): {f.error}");
            // 列出生成的文件
            var outFiles = Directory.GetFiles(outputDir, "*.xlsx", SearchOption.AllDirectories);
            Log($"  输出目录文件数: {outFiles.Length}");
            foreach (var f in outFiles)
                Log($"    {Path.GetRelativePath(outputDir, f)} ({new FileInfo(f).Length} bytes)");

            // 6. 验证生成结果：打开一个输出文件检查内容
            if (outFiles.Length > 0)
            {
                Log("\n--- 6. 验证生成结果 ---");
                using var wb = new ClosedXML.Excel.XLWorkbook(outFiles[0]);
                var ws = wb.Worksheets.First();
                Log($"  打开文件: {Path.GetFileName(outFiles[0])}");
                Log($"  Sheet: {ws.Name}");
                for (int r = 1; r <= Math.Min(6, ws.LastRowUsed()?.RowNumber() ?? 0); r++)
                {
                    var vals = new List<string>();
                    for (int c = 1; c <= Math.Min(4, ws.LastColumnUsed()?.ColumnNumber() ?? 0); c++)
                        vals.Add(ws.Cell(r, c).Value.ToString() ?? "");
                    Log($"  行{r}: {string.Join(" | ", vals)}");
                }
                // 检查公式是否转为值
                var formulaCell = ws.Cell(6, 2);
                Log($"  B6值: {formulaCell.Value} (公式:{formulaCell.FormulaA1})");
            }

            Log("\n========== 冒烟测试完成 ==========");
        }
        catch (Exception ex)
        {
            Log($"ERROR: {ex.GetType().FullName}: {ex.Message}");
            log.AppendLine(ex.StackTrace ?? "");
            log.AppendLine($"Inner: {ex.InnerException?.Message}");
        }
        finally
        {
            File.WriteAllText(logFile, log.ToString());
        }
    }

    private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
    {
        var logPath = Path.Combine(Path.GetTempPath(), "ExcelMailMerge_runtime_error.txt");
        File.WriteAllText(logPath, FlattenException(e.Exception));
        MessageBox.Show($"{e.Exception.GetType().Name}: {e.Exception.Message}\n\n详情已写入: {logPath}", "运行时错误", MessageBoxButton.OK, MessageBoxImage.Error);
        e.Handled = true;
    }

    /// <summary>
    /// 递归展开异常（含 TypeInitializationException 的多层 InnerException）
    /// </summary>
    private static string FlattenException(Exception? ex)
    {
        if (ex is null) return "（无异常）";
        var sb = new StringBuilder();
        int depth = 0;
        var current = ex;
        while (current != null && depth < 10)
        {
            sb.AppendLine($"===== [Level {depth}] {current.GetType().FullName} =====");
            sb.AppendLine($"Message: {current.Message}");
            if (current.Source != null)
                sb.AppendLine($"Source: {current.Source}");
            sb.AppendLine("StackTrace:");
            sb.AppendLine(current.StackTrace ?? "（无堆栈）");
            sb.AppendLine();
            current = current.InnerException;
            depth++;
        }
        return sb.ToString();
    }
}
