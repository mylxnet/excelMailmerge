using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelMailMerge.Services;

/// <summary>
/// 文件格式转换器：通过 COM 互操作调用 WPS 或 Excel，将 .xls/.et 转换为 .xlsx
/// </summary>
public static class FileFormatConverter
{
    private static readonly string[] SupportedNonXlsxExtensions = { ".xls", ".et" };
    private static readonly string TempDir = Path.Combine(Path.GetTempPath(), "ExcelMailMerge_converted");

    /// <summary>
    /// 确保文件为 .xlsx 格式。如果已是 .xlsx 则原样返回；
    /// 如果是 .xls 或 .et 则通过 COM 互操作转换为 .xlsx。
    /// </summary>
    public static string EnsureXlsx(string filePath)
    {
        if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            return filePath;

        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        if (ext is ".xlsx" or ".xlsm")
            return filePath;

        if (!SupportedNonXlsxExtensions.Contains(ext))
            return filePath;

        Directory.CreateDirectory(TempDir);
        var baseName = Path.GetFileNameWithoutExtension(filePath);
        var tempPath = Path.Combine(TempDir, $"{baseName}_{DateTime.Now:yyyyMMddHHmmss}.xlsx");

        try
        {
            ConvertToXlsxViaCom(filePath, tempPath);
            return tempPath;
        }
        catch (Exception ex)
        {
            try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { }
            throw new InvalidOperationException(
                $"无法转换文件 '{Path.GetFileName(filePath)}' 为 .xlsx 格式。\n" +
                $"原因：{ex.Message}\n\n请确认已安装 WPS 或 Microsoft Excel。", ex);
        }
    }

    /// <summary>
    /// 通过 COM 互操作打开源文件并另存为 .xlsx 格式
    /// </summary>
    private static void ConvertToXlsxViaCom(string source, string dest)
    {
        Type? appType = null;
        object? app = null;

        try
        {
            // 依次尝试 WPS 表格 → WPS备用ProgID → Microsoft Excel
            appType = Type.GetTypeFromProgID("et.Application")
                   ?? Type.GetTypeFromProgID("Ket.Application")
                   ?? Type.GetTypeFromProgID("Excel.Application");

            if (appType == null)
                throw new InvalidOperationException("未检测到 WPS 或 Microsoft Excel。");

            app = Activator.CreateInstance(appType);

            // 设置不可见、不弹警告
            appType.InvokeMember("Visible", BindingFlags.SetProperty, null, app, new object[] { false });
            try { appType.InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, app, new object[] { false }); } catch { }

            // 获取 Workbooks 集合
            var workbooks = appType.InvokeMember("Workbooks", BindingFlags.GetProperty, null, app, null);

            // 打开源文件
            var wb = workbooks!.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, workbooks,
                new object[] { source });

            // 另存为 .xlsx（51 = xlOpenXMLWorkbook）
            // SaveAs 有多个可选参数，用 Missing.Value 填充
            var missing = Missing.Value;
            wb!.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, wb,
                new object[] { dest, 51, missing, missing, missing, missing,
                               missing, missing, missing, missing, missing, missing });

            // 关闭工作簿
            wb.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, wb, new object[] { false });

            // 释放 COM 对象
            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(workbooks);
        }
        finally
        {
            if (app != null)
            {
                try { appType!.InvokeMember("Quit", BindingFlags.InvokeMethod, null, app, null); } catch { }
                try { Marshal.ReleaseComObject(app); } catch { }
            }
        }
    }

    /// <summary>清理所有转换产生的临时文件</summary>
    public static void Cleanup()
    {
        try { if (Directory.Exists(TempDir)) Directory.Delete(TempDir, recursive: true); } catch { }
    }
}
