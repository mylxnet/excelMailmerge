using ClosedXML.Excel;
using ExcelMailMerge.Models;

namespace ExcelMailMerge.Services;

/// <summary>
/// 数据源解析服务：读取Excel、选Sheet、解析列名、读取数据行
/// </summary>
public class DataSourceService
{
    /// <summary>
    /// 获取文件中所有Sheet的信息
    /// </summary>
    public List<DataSourceSheetInfo> GetSheetInfos(string filePath)
    {
        var result = new List<DataSourceSheetInfo>();
        using var wb = new XLWorkbook(filePath);
        foreach (var ws in wb.Worksheets)
        {
            var range = ws.RangeUsed();
            var rowCount = range?.RowCount() ?? 0;
            var colCount = range?.ColumnCount() ?? 0;
            result.Add(new DataSourceSheetInfo
            {
                SheetName = ws.Name,
                RowCount = rowCount,
                ColumnCount = colCount,
                HasData = rowCount > 0 && colCount > 0
            });
        }
        return result;
    }

    /// <summary>
    /// 解析数据源：列名 + 所有数据行
    /// </summary>
    /// <param name="filePath">数据源文件路径</param>
    /// <param name="sheetName">Sheet名</param>
    /// <param name="titleRowIndex">标题行号（1-based）</param>
    public ParsedDataSource Parse(string filePath, string sheetName, int titleRowIndex)
    {
        var result = new ParsedDataSource
        {
            FilePath = filePath,
            SheetName = sheetName,
            TitleRowIndex = titleRowIndex,
            DataStartRowIndex = titleRowIndex + 1
        };

        using var wb = new XLWorkbook(filePath);
        var ws = wb.Worksheet(sheetName);
        var usedRange = ws.RangeUsed();
        if (usedRange == null) return result;

        var lastCol = usedRange.LastColumn().ColumnNumber();
        var lastRow = usedRange.LastRow().RowNumber();

        // 1. 解析列名
        var columns = new List<ColumnInfo>();
        for (int col = 1; col <= lastCol; col++)
        {
            var cell = ws.Cell(titleRowIndex, col);
            var rawVal = cell.GetString();
            if (string.IsNullOrWhiteSpace(rawVal)) continue; // 空列跳过
            columns.Add(new ColumnInfo
            {
                OriginalName = rawVal,
                ColumnIndex = col
            });
        }
        result.Columns = columns;

        // 2. 检查标题列是否含大括号（非法）
        foreach (var col in columns)
        {
            if (col.DisplayName.Contains('{') || col.DisplayName.Contains('}'))
            {
                throw new InvalidOperationException(
                    $"标题列 [{col.DisplayName}] 中包含大括号字符 '{{' 或 '}}'，请修改后重新上传。" +
                    $"大括号仅用于模板中标识占位符，不能作为列名出现。");
            }
        }

        // 3. 检查数据区域是否有合并单元格（非法）
        if (ws.MergedRanges.Any(r =>
                r.FirstRow().RowNumber() >= result.DataStartRowIndex &&
                r.FirstRow().RowNumber() <= lastRow))
        {
            throw new InvalidOperationException(
                "数据区域检测到合并单元格！为避免读取错位，数据区域不允许有合并单元格。" +
                "请先在Excel中取消合并、补齐数据后再上传。");
        }

        // 4. 读取数据行，遇到整行全空停止
        int displayIndex = 1;
        for (int row = result.DataStartRowIndex; row <= lastRow; row++)
        {
            // 判断是否整行空
            bool isRowEmpty = true;
            var values = new Dictionary<string, object?>();
            foreach (var col in columns)
            {
                var cell = ws.Cell(row, col.ColumnIndex);
                var val = ReadCellValue(cell);
                values[col.DisplayName] = val;
                if (val != null && !string.IsNullOrWhiteSpace(val.ToString()))
                {
                    isRowEmpty = false;
                }
            }
            if (isRowEmpty) break; // 整行空→停止

            result.Rows.Add(new DataRow
            {
                OriginalRowIndex = row,
                DisplayIndex = displayIndex++,
                Values = values
            });
        }

        return result;
    }

    /// <summary>
    /// 读取单元格值：保留原始类型（DateTime、数字、字符串等）
    /// </summary>
    private static object? ReadCellValue(IXLCell cell)
    {
        if (cell == null || cell.Value.IsBlank) return null;
        var val = cell.Value;
        if (val.IsText) return val.GetText();
        if (val.IsNumber) return val.GetNumber();
        if (val.IsBoolean) return val.GetBoolean();
        if (val.IsDateTime) return val.GetDateTime();
        if (val.IsTimeSpan) return val.GetTimeSpan();
        // 其他情况（错误值等）统一转字符串
        return cell.GetString();
    }

    /// <summary>
    /// 自动检测标题行号：从第1行开始向下找第一个非空行
    /// </summary>
    public int AutoDetectTitleRow(string filePath, string sheetName)
    {
        using var wb = new XLWorkbook(filePath);
        var ws = wb.Worksheet(sheetName);
        var usedRange = ws.RangeUsed();
        if (usedRange == null) return 1;
        var lastRow = usedRange.LastRow().RowNumber();
        var lastCol = usedRange.LastColumn().ColumnNumber();

        for (int row = 1; row <= Math.Min(lastRow, 20); row++)
        {
            bool hasAnyValue = false;
            int nonEmptyCount = 0;
            for (int col = 1; col <= lastCol; col++)
            {
                var cell = ws.Cell(row, col);
                var txt = cell.GetString();
                if (!string.IsNullOrWhiteSpace(txt))
                {
                    hasAnyValue = true;
                    nonEmptyCount++;
                }
            }
            if (hasAnyValue && nonEmptyCount >= Math.Max(1, lastCol / 2))
                return row; // 至少有一半列有内容，认为是标题行
        }
        return 1; // 兜底
    }
}
