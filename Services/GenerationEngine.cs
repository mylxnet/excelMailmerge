using System.IO;
using System.Threading;
using ClosedXML.Excel;
using ExcelMailMerge.Models;

namespace ExcelMailMerge.Services;

/// <summary>
/// 核心生成引擎：执行占位符替换 + 生成最终输出
/// </summary>
public class GenerationEngine
{
    private readonly NamingService _naming;
    private readonly TemplateService _tplSvc;

    public GenerationEngine(NamingService naming, TemplateService tplSvc) { _naming = naming; _tplSvc = tplSvc; }

    // 磁盘追踪日志（独立于 UI Progress，即使挂住也能看到进度）
    private static readonly string TraceFile = Path.Combine(Path.GetTempPath(), "ExcelMailMerge_trace.log");
    private static void Trace(string msg)
    {
        try { File.AppendAllText(TraceFile, $"[{DateTime.Now:HH:mm:ss.fff}] {msg}\n"); } catch { }
    }

    /// <summary>
    /// 执行生成（异步可取消，报告进度）
    /// </summary>
    public async Task<GenerationProgress> GenerateAsync(
        ParsedDataSource dataSource,
        ScannedTemplate template,
        List<string> namingColumns,
        string namingSeparator,
        OutputMode outputMode,
        string? customOutputFolder,
        AdvancedWriteSettings writeSettings,
        string? singleFileName,
        IProgress<GenerationProgress>? progress = null,
        CancellationToken ct = default)
    {
        var pg = new GenerationProgress { TotalRows = dataSource.Rows.Count };

        // 1. 确定输出根目录，强制新建时间戳子文件夹
        var rootFolder = string.IsNullOrWhiteSpace(customOutputFolder)
            ? Path.GetDirectoryName(dataSource.FilePath) ?? Directory.GetCurrentDirectory()
            : customOutputFolder;
        var timestampFolderName = $"输出结果{DateTime.Now:yyyyMMdd_HHmmss}";
        var outputFolder = Path.Combine(rootFolder, timestampFolderName);
        Directory.CreateDirectory(outputFolder);
        pg.OutputFolder = outputFolder;

        var namingResults = _naming.ValidateAllNames(dataSource, namingColumns, namingSeparator).perRowResults;

        try
        {
            // 切到线程池执行，避免 ClosedXML 的 CPU 密集型操作阻塞 UI 线程
            await Task.Run(async () =>
            {
                if (outputMode == OutputMode.MultiFilePerRow)
                    await RunModeB(dataSource, template, namingResults, outputFolder, writeSettings, pg, progress, ct);
                else
                    await RunModeA(dataSource, template, namingResults, namingSeparator, outputFolder, writeSettings,
                        singleFileName, pg, progress, ct);
            });
        }
        catch (OperationCanceledException)
        {
            pg.IsCanceled = true;
        }
        finally
        {
            pg.IsCompleted = true;
            progress?.Report(pg);
            ExportLog(pg);
        }

        // 自动打开输出目录的逻辑移到 MainViewModel（避免污染命令行冒烟测试）
        return pg;
    }

    /// <summary>
    /// 试生成：用第1条数据填充模板，生成临时预览文件（不进正式输出目录）
    /// </summary>
    public PreviewResult GeneratePreview(
        ParsedDataSource dataSource,
        ScannedTemplate template,
        AdvancedWriteSettings writeSettings)
    {
        var result = new PreviewResult();
        if (dataSource.Rows.Count == 0)
        {
            result.ErrorMessage = "数据源无有效数据行";
            return result;
        }

        var firstRow = dataSource.Rows[0];
        result.SourceRowIndex = firstRow.OriginalRowIndex;
        result.TempFileName = $"预览_第{firstRow.OriginalRowIndex}行.xlsx";

        try
        {
            // 临时文件放到系统临时目录，避免污染正式输出目录
            var tempDir = Path.Combine(Path.GetTempPath(), "ExcelMailMerge_Preview");
            Directory.CreateDirectory(tempDir);
            var tempPath = Path.Combine(tempDir, result.TempFileName);
            // 若已存在先删，保证最新
            try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { /* 忽略 */ }

            using var wb = new XLWorkbook(template.FilePath);
            int placeholderReplaced = 0;

            // 第一遍：替换占位符
            foreach (var ws in wb.Worksheets)
            {
                FillWorksheetWithStats(ws, firstRow, writeSettings, ref placeholderReplaced);
            }

            // 第二遍：把所有公式转为计算值（需求：公式不保留）
            // 必须在占位符替换完成后进行，这样 cell.Value 才能基于最新值计算公式
            int formulaConverted = ConvertFormulasToValues(wb, writeSettings);

            wb.SaveAs(tempPath);

            result.IsSuccess = true;
            result.TempFilePath = tempPath;
            result.SheetCount = wb.Worksheets.Count;
            result.PlaceholderReplacedCount = placeholderReplaced;
            result.FormulaConvertedCount = formulaConverted;
            result.PlaceholderUnmatchedCount = template.UnmatchedCount;

            // 读取临时文件内容作为填充预览
            var (filledRows, filledCols) = _tplSvc.GetTemplateContent(tempPath);
            result.FilledContentRows = filledRows;
            result.FilledContentColumns = filledCols;

            return result;
        }
        catch (Exception ex)
        {
            result.ErrorMessage = ex.Message;
            return result;
        }
    }

    /// <summary>
    /// 清除试生成的临时预览文件
    /// </summary>
    public void ClearPreviewFile(string? tempFilePath)
    {
        if (string.IsNullOrEmpty(tempFilePath) || !File.Exists(tempFilePath)) return;
        try { File.Delete(tempFilePath); } catch { /* 忽略删除失败 */ }
    }

    /// <summary>
    /// 带统计的单元格填充（用于试生成）
    /// 注意：这里只替换占位符，公式→值的转换由 ConvertFormulasToValues 统一处理
    /// </summary>
    private static void FillWorksheetWithStats(IXLWorksheet ws, DataRow row,
        AdvancedWriteSettings writeSettings, ref int placeholderReplaced)
    {
        var used = ws.RangeUsed();
        if (used == null) return;
        foreach (var cell in used.Cells())
        {
            string? original;
            bool hadFormula = false;
            if (cell.HasFormula)
            {
                // 公式单元格：占位符替换基于缓存值（公式本身不会包含{占位符}，但其引用的单元格可能已替换）
                hadFormula = true;
                try { var v = cell.CachedValue; original = string.IsNullOrEmpty(v.ToString()) ? null : v.ToString(); }
                catch { original = null; }
            }
            else
            {
                original = cell.GetString();
            }
            if (string.IsNullOrEmpty(original)) continue;

            var (newValue, hasPh) = TemplateService.ReplacePlaceholders(original, row, writeSettings, outputType: 1);
            if (!hasPh) continue;

            placeholderReplaced++;

            if (newValue == null || newValue is DBNull)
            {
                cell.Value = Blank.Value;
                continue;
            }
            if (hadFormula) cell.Clear(XLClearOptions.AllContents);
            WriteCellWithType(cell, newValue, writeSettings);
        }
    }

    /// <summary>
    /// 把工作簿内所有公式转换为字面值（需求：公式不保留，直接输入计算值）
    /// 必须在占位符替换完成后调用，这样 cell.Value 才能基于最新值重新计算公式
    /// </summary>
    private static int ConvertFormulasToValues(IXLWorkbook wb, AdvancedWriteSettings writeSettings)
    {
        int converted = 0;
        foreach (var ws in wb.Worksheets)
        {
            var used = ws.RangeUsed();
            if (used == null) continue;
            foreach (var cell in used.Cells())
            {
                if (!cell.HasFormula) continue;
                try
                {
                    // cell.Value 会基于当前单元格的值重新计算公式（例如 =B5*D5 用替换后的 B5/D5 计算）
                    XLCellValue computedValue = cell.Value;
                    // 清掉公式（保留样式，XLClearOptions.AllContents 只清内容不清格式）
                    cell.Clear(XLClearOptions.AllContents);
                    // 直接赋值 XLCellValue，保留原始类型（数字/文本/日期/布尔）
                    cell.Value = computedValue;

                    // 长数字转文本处理（针对计算结果是长数字的情况）
                    if (writeSettings.AutoLongNumberAsText)
                    {
                        var numStr = computedValue.ToString();
                        if (TemplateService.IsLongPureNumber(numStr))
                        {
                            cell.Style.NumberFormat.Format = "@";
                            cell.SetValue(numStr);
                        }
                    }
                    converted++;
                }
                catch { /* 忽略单个公式单元格转换失败，继续处理其他 */ }
            }
        }
        return converted;
    }

    // ========== 模式B：每行→独立文件 ==========
    private Task RunModeB(
        ParsedDataSource dataSource,
        ScannedTemplate template,
        List<NamingResult> namingResults,
        string outputFolder,
        AdvancedWriteSettings writeSettings,
        GenerationProgress pg,
        IProgress<GenerationProgress>? progress,
        CancellationToken ct)
    {
        for (int i = 0; i < dataSource.Rows.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            var row = dataSource.Rows[i];
            var nr = namingResults[i];
            pg.CurrentRowIndex = i + 1;

            try
            {
                if (!nr.IsValid)
                    throw new InvalidOperationException($"命名组合无效：{nr.ErrorMessage}");

                // 1. 打开模板，复制整套WB（不读入内存，ClosedXML直接SaveAs新文件）
                using var wb = new XLWorkbook(template.FilePath);

                // 2. 遍历所有Sheet，逐单元格替换占位符
                foreach (var ws in wb.Worksheets)
                    FillWorksheet(ws, row, writeSettings);

                // 3. 公式→值（必须在占位符替换后，这样公式能用新值重新计算）
                ConvertFormulasToValues(wb, writeSettings);

                // 4. 生成唯一文件名并保存
                var baseName = _naming.Truncate(_naming.SanitizeForFileName(nr.Name), NamingService.MaxFileNameLength);
                var outPath = _naming.EnsureUniqueFileName(outputFolder, baseName, "xlsx");
                wb.SaveAs(outPath);

                pg.SuccessCount++;
                pg.Logs.Add($"[OK]   第{row.OriginalRowIndex}行 → {Path.GetFileName(outPath)}");
            }
            catch (Exception ex)
            {
                pg.FailCount++;
                pg.Failures.Add((row.OriginalRowIndex, nr.Name, ex.Message));
                pg.Logs.Add($"[FAIL] 第{row.OriginalRowIndex}行 [{nr.Name}] → {ex.Message}");
            }

            // 节流：每5行或最后一行才报告进度
            if (i % 5 == 0 || i == dataSource.Rows.Count - 1)
                progress?.Report(pg);
        }
        return Task.CompletedTask;
    }

    // ========== 模式A：所有行→同一文件 ==========
    private Task RunModeA(
        ParsedDataSource dataSource,
        ScannedTemplate template,
        List<NamingResult> namingResults,
        string namingSeparator,
        string outputFolder,
        AdvancedWriteSettings writeSettings,
        string? singleFileName,
        GenerationProgress pg,
        IProgress<GenerationProgress>? progress,
        CancellationToken ct)
    {
        // 安全检查：Sheet总数不得超过Excel上限
        const int MaxSheetsPerWorkbook = 200;
        int totalSheets = template.SheetCount * dataSource.Rows.Count;
        if (totalSheets > MaxSheetsPerWorkbook)
        {
            pg.FailCount = dataSource.Rows.Count;
            pg.Logs.Add($"[ERROR] 单文件模式Sheet总数 {totalSheets} 超过上限 {MaxSheetsPerWorkbook}，已阻止生成。");
            return Task.CompletedTask;
        }

        var outFileName = string.IsNullOrWhiteSpace(singleFileName)
            ? $"合并结果_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
            : (_naming.SanitizeForFileName(singleFileName!.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                ? singleFileName[..^5]
                : singleFileName) + ".xlsx");
        var outPath = Path.Combine(outputFolder, outFileName);

        using var outputWb = new XLWorkbook();
        var usedSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // 模板只打开一次，循环内复用（避免每行重新解析模板文件）
        var tplSw = System.Diagnostics.Stopwatch.StartNew();
        using var tplWb = new XLWorkbook(template.FilePath);
        tplSw.Stop();
        pg.Logs.Add($"[计时] 打开模板: {tplSw.ElapsedMilliseconds}ms");
        Trace($"RunModeA 开始: {dataSource.Rows.Count}行, 模板Sheet数={template.SheetCount}");

        var loopSw = System.Diagnostics.Stopwatch.StartNew();
        for (int i = 0; i < dataSource.Rows.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            var row = dataSource.Rows[i];
            var nr = namingResults[i];
            pg.CurrentRowIndex = i + 1;

            try
            {
                if (!nr.IsValid)
                    throw new InvalidOperationException($"命名组合无效：{nr.ErrorMessage}");

                var baseName = _naming.SanitizeForFileName(nr.Name);

                foreach (var tplWs in tplWb.Worksheets)
                {
                    var newSheetName = _naming.EnsureUniqueSheetName(usedSheetNames, baseName, tplWs.Name);
                    var destWs = outputWb.AddWorksheet(newSheetName);
                    CopySheetContentAndStyle(tplWs, destWs);
                    FillWorksheet(destWs, row, writeSettings);
                }

                pg.SuccessCount++;
                // 只记录前5行和每20行的日志，避免日志过多
                if (i < 5 || i % 20 == 0)
                    pg.Logs.Add($"[OK]   第{row.OriginalRowIndex}行（{baseName}）→ Sheet已加入");
            }
            catch (Exception ex)
            {
                pg.FailCount++;
                pg.Failures.Add((row.OriginalRowIndex, nr.Name, ex.Message));
                pg.Logs.Add($"[FAIL] 第{row.OriginalRowIndex}行 [{nr.Name}] → {ex.Message}");
            }

            // 磁盘追踪：每10行写一次
            if (i % 10 == 0)
                Trace($"  行 {i+1}/{dataSource.Rows.Count} 完成, 成功={pg.SuccessCount}, 耗时={loopSw.ElapsedMilliseconds}ms");

            // 节流：每5行或最后一行才报告进度，避免大量 Dispatcher Post 淹没 UI 线程
            if (i % 5 == 0 || i == dataSource.Rows.Count - 1)
                progress?.Report(pg);
        }
        loopSw.Stop();
        pg.Logs.Add($"[计时] 循环处理 {dataSource.Rows.Count} 行: {loopSw.ElapsedMilliseconds}ms");
        Trace($"循环完成: {loopSw.ElapsedMilliseconds}ms, 开始公式转值...");
        pg.Logs.Add("[进度] 正在转换公式为值...");
        progress?.Report(pg);

        var formulaSw = System.Diagnostics.Stopwatch.StartNew();
        ConvertFormulasToValues(outputWb, writeSettings);
        formulaSw.Stop();
        Trace($"公式转值完成: {formulaSw.ElapsedMilliseconds}ms, 开始保存...");
        pg.Logs.Add($"[计时] 公式转值: {formulaSw.ElapsedMilliseconds}ms");
        pg.Logs.Add("[进度] 正在保存文件（单文件模式Sheet较多时可能较慢）...");
        progress?.Report(pg);

        var saveSw = System.Diagnostics.Stopwatch.StartNew();
        outputWb.SaveAs(outPath);
        saveSw.Stop();
        Trace($"保存完成: {saveSw.ElapsedMilliseconds}ms, 总结束");
        pg.Logs.Add($"[计时] 保存文件: {saveSw.ElapsedMilliseconds}ms");
        pg.Logs.Add($"[计时] 总耗时: {tplSw.ElapsedMilliseconds + loopSw.ElapsedMilliseconds + formulaSw.ElapsedMilliseconds + saveSw.ElapsedMilliseconds}ms");
        progress?.Report(pg);

        return Task.CompletedTask;
    }

    // ========== 工具方法：逐单元格填充 ==========
    private static void FillWorksheet(IXLWorksheet ws, DataRow row, AdvancedWriteSettings writeSettings)
    {
        var used = ws.RangeUsed();
        if (used == null) return;
        foreach (var cell in used.Cells())
        {
            // 1. 读取原始内容：公式→取CachedValue；文本→直接取
            string? original;
            bool hadFormula = false;
            if (cell.HasFormula)
            {
                hadFormula = true;
                try { var v = cell.CachedValue; original = string.IsNullOrEmpty(v.ToString()) ? null : v.ToString(); }
                catch { original = null; }
            }
            else
            {
                original = cell.GetString();
            }
            if (string.IsNullOrEmpty(original)) continue;
            // 2. 替换占位符
            var (newValue, hasPh) = TemplateService.ReplacePlaceholders(original, row, writeSettings, outputType: 1);
            if (!hasPh) continue; // 没占位符，跳过（保留原样）

            // 3. 写入新值
            if (newValue == null || newValue is DBNull)
            {
                cell.Value = Blank.Value; // 真正清空
                continue;
            }

            // 如果原来有公式，现在已经是纯值了，需要清掉公式并重新设值
            if (hadFormula) cell.Clear(XLClearOptions.AllContents);

            // 写入并处理：长数字强制文本、DateTime保留类型等
            WriteCellWithType(cell, newValue, writeSettings);
        }
    }

    private static void WriteCellWithType(IXLCell cell, object value, AdvancedWriteSettings settings)
    {
        switch (value)
        {
            case string s:
                if (settings.AutoLongNumberAsText && TemplateService.IsLongPureNumber(s))
                {
                    // 强制文本格式：设DataType + 前面加单引号风格
                    cell.Style.NumberFormat.Format = "@";
                    cell.Value = s;
                    cell.SetValue(s);
                }
                else
                {
                    cell.SetValue(s);
                }
                break;
            case DateTime dt:
                cell.SetValue(dt);
                break;
            case double d:
                cell.SetValue(d);
                break;
            case decimal dec:
                cell.SetValue(dec);
                break;
            case int i:
                cell.SetValue(i);
                break;
            case long l:
                // long可能超Excel精度（>15位）→ 若AutoLongNumberAsText且长度>11则按文本
                var lstr = l.ToString();
                if (settings.AutoLongNumberAsText && lstr.Length > 11)
                {
                    cell.Style.NumberFormat.Format = "@";
                    cell.SetValue(lstr);
                }
                else cell.SetValue(l);
                break;
            case bool b:
                cell.SetValue(b);
                break;
            default:
                cell.SetValue(value.ToString() ?? string.Empty);
                break;
        }
    }

    // ========== 工具方法：跨工作簿复制Sheet内容+样式 ==========
    private static void CopySheetContentAndStyle(IXLWorksheet src, IXLWorksheet dst)
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();
        var srcUsed = src.RangeUsed();
        if (srcUsed == null) return;

        int firstCol = srcUsed.FirstColumn().ColumnNumber();
        int lastCol = srcUsed.LastColumn().ColumnNumber();
        int firstRow = srcUsed.FirstRow().RowNumber();
        int lastRow = srcUsed.LastRow().RowNumber();
        Trace($"    CopySheet: RangeUsed firstCol={firstCol} lastCol={lastCol} firstRow={firstRow} lastRow={lastRow} [{sw.ElapsedMilliseconds}ms]");

        // 扩展1列1行：确保最右列右边框、最末行下边框也被覆盖
        lastCol++;
        lastRow++;

        // 进一步扩展列范围：覆盖无内容但有自定义宽度的列
        int maxColProbe = lastCol + 30;
        for (int c = lastCol + 1; c <= maxColProbe; c++)
        {
            try
            {
                if (Math.Abs(src.Column(c).Width - src.ColumnWidth) > 0.01)
                    lastCol = c;
            }
            catch { break; }
        }
        Trace($"    CopySheet: 列探查完成 lastCol={lastCol} [{sw.ElapsedMilliseconds}ms]");

        // 进一步扩展行范围：覆盖无内容但有自定义行高的行
        // 注意：必须比较自定义行高与默认行高，否则默认行高15 > 0恒为true导致无限循环
        double defaultRowHeight = src.RowHeight;
        int maxRowProbe = lastRow + 30;
        for (int r = lastRow + 1; r <= maxRowProbe; r++)
        {
            try
            {
                if (Math.Abs(src.Row(r).Height - defaultRowHeight) > 0.01)
                    lastRow = r;
            }
            catch { break; }
        }
        Trace($"    CopySheet: 行探查完成 lastRow={lastRow} [{sw.ElapsedMilliseconds}ms]");

        var srcRange = src.Range(firstRow, firstCol, lastRow, lastCol);

        // 1. 一次性 CopyTo：值+样式+合并区域+数字格式+条件格式，全都拷贝
        srcRange.CopyTo(dst.FirstCell());
        Trace($"    CopySheet: CopyTo完成 ({lastRow-firstRow+1}行×{lastCol-firstCol+1}列) [{sw.ElapsedMilliseconds}ms]");

        // 2. 先设置工作表级默认值（列宽/行高/Tab颜色）
        try { dst.ColumnWidth = src.ColumnWidth; } catch { }
        try { dst.RowHeight = src.RowHeight; } catch { }
        try { dst.TabColor = src.TabColor; } catch { }

        // 3. 复制所有列宽（从第1列开始）
        for (int c = 1; c <= lastCol; c++)
        {
            try
            {
                dst.Column(c).Width = src.Column(c).Width;
                if (src.Column(c).IsHidden) dst.Column(c).Hide();
            }
            catch { }
        }
        Trace($"    CopySheet: 列宽复制完成 ({lastCol}列) [{sw.ElapsedMilliseconds}ms]");

        // 4. 复制所有行高（从第1行开始）
        for (int r = 1; r <= lastRow; r++)
        {
            try
            {
                if (src.Row(r).Height > 0)
                    dst.Row(r).Height = src.Row(r).Height;
                if (src.Row(r).IsHidden) dst.Row(r).Hide();
            }
            catch { }
        }
        Trace($"    CopySheet: 行高复制完成 ({lastRow}行) [{sw.ElapsedMilliseconds}ms]");
    }

    // ========== 日志导出 ==========
    private static void ExportLog(GenerationProgress pg)
    {
        try
        {
            var logPath = Path.Combine(pg.OutputFolder, $"生成日志_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
            var lines = new List<string>
            {
                "======== Excel邮件合并 生成日志 ========",
                $"生成时间：{DateTime.Now:yyyy-MM-dd HH:mm:ss}",
                $"总数据行数：{pg.TotalRows}",
                $"成功：{pg.SuccessCount}  失败：{pg.FailCount}",
                $"状态：{pg.StatusText}",
                $"输出目录：{pg.OutputFolder}",
                "",
                "-------- 详细记录 --------"
            };
            lines.AddRange(pg.Logs);
            if (pg.Failures.Count > 0)
            {
                lines.Add("");
                lines.Add("-------- 失败明细 --------");
                foreach (var (row, name, err) in pg.Failures)
                    lines.Add($"第{row}行 [{name}]：{err}");
            }
            File.WriteAllLines(logPath, lines);
        }
        catch { /* 忽略日志导出失败 */ }
    }
}
