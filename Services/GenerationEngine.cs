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
            if (outputMode == OutputMode.MultiFilePerRow)
                await RunModeB(dataSource, template, namingResults, outputFolder, writeSettings, pg, progress, ct);
            else
                await RunModeA(dataSource, template, namingResults, namingSeparator, outputFolder, writeSettings,
                    singleFileName, pg, progress, ct);
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
    private async Task RunModeB(
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

            progress?.Report(pg);
            // 每20条让出一点UI时间
            if (i % 20 == 0) await Task.Yield();
        }
    }

    // ========== 模式A：所有行→同一文件 ==========
    private async Task RunModeA(
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
        var outFileName = string.IsNullOrWhiteSpace(singleFileName)
            ? $"合并结果_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
            : (_naming.SanitizeForFileName(singleFileName!.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                ? singleFileName[..^5]
                : singleFileName) + ".xlsx");
        var outPath = Path.Combine(outputFolder, outFileName);

        using var outputWb = new XLWorkbook();
        var usedSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // 先把模板所有Sheet作为"母版"复制一遍，然后在内存里复制
        // 更高效的做法是：每一行 → 对模板每个Sheet做CloneTo，然后填入
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

                // 打开模板只读一遍，逐Sheet克隆
                using var tplWb = new XLWorkbook(template.FilePath);
                var baseName = _naming.SanitizeForFileName(nr.Name);

                foreach (var tplWs in tplWb.Worksheets)
                {
                    // 生成新Sheet名（加命名前缀，防重名）
                    var newSheetName = _naming.EnsureUniqueSheetName(usedSheetNames, baseName, tplWs.Name);

                    // 复制：ClosedXML无直接跨WB的Sheet克隆，所以把整个模板工作簿当作源，逐单元格复制
                    // 简单方案：将tplWs的UsedRange复制到新的目标Sheet
                    var destWs = outputWb.AddWorksheet(newSheetName);
                    CopySheetContentAndStyle(tplWs, destWs);

                    // 填充占位符
                    FillWorksheet(destWs, row, writeSettings);
                }

                pg.SuccessCount++;
                pg.Logs.Add($"[OK]   第{row.OriginalRowIndex}行（{baseName}）→ Sheet已加入");
            }
            catch (Exception ex)
            {
                pg.FailCount++;
                pg.Failures.Add((row.OriginalRowIndex, nr.Name, ex.Message));
                pg.Logs.Add($"[FAIL] 第{row.OriginalRowIndex}行 [{nr.Name}] → {ex.Message}");
            }

            progress?.Report(pg);
            if (i % 10 == 0) await Task.Yield();
        }

        // 模式A：所有占位符替换完成后，统一把所有公式转为值
        ConvertFormulasToValues(outputWb, writeSettings);
        outputWb.SaveAs(outPath);
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
        var srcUsed = src.RangeUsed();
        if (srcUsed == null) return;

        int firstCol = srcUsed.FirstColumn().ColumnNumber();
        int lastCol = srcUsed.LastColumn().ColumnNumber();
        int firstRow = srcUsed.FirstRow().RowNumber();
        int lastRow = srcUsed.LastRow().RowNumber();
        // 扩展1列1行：确保最右列右边框、最末行下边框也被覆盖（RangeUsed只算"有内容"的单元格）
        lastCol++;
        lastRow++;

        var srcRange = src.Range(firstRow, firstCol, lastRow, lastCol);

        // 1. 一次性 CopyTo：值+样式+合并区域+数字格式+条件格式，全都拷贝
        srcRange.CopyTo(dst.FirstCell());

        // 2. 单独复制列宽和行高（CopyTo 不带这个）
        for (int c = firstCol; c <= lastCol; c++)
        {
            try { dst.Column(c).Width = src.Column(c).Width; } catch { }
        }
        for (int r = firstRow; r <= lastRow; r++)
        {
            try { dst.Row(r).Height = src.Row(r).Height; } catch { }
        }

        // 3. Sheet 属性
        try { dst.TabColor = src.TabColor; } catch { }
        try { dst.ColumnWidth = src.ColumnWidth; } catch { }
        try { dst.RowHeight = src.RowHeight; } catch { }

        // 4. 保留 dst 的 UsedRange（不含空填充列）给后续填充占位符用
        //    ClosedXML 的 CopyTo 已把值/样式/合并复制完毕，UsedRange 扫描会自动识别
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
