using System.IO;
using ExcelMailMerge.Models;
using Newtonsoft.Json;

namespace ExcelMailMerge.Services;

/// <summary>
/// 配置持久化服务：读写 AppData 下的 settings.json
/// </summary>
public class SettingsService
{
    private readonly string _settingsFolder;
    private readonly string _settingsFilePath;

    public SettingsService()
    {
        // 优先使用 AppData\Local\ExcelMailMerge；若不可用（权限/锁定），依次降级到 程序目录、临时目录
        var candidates = new[]
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelMailMerge"),
            Path.Combine(AppContext.BaseDirectory, "config"),
            Path.Combine(Path.GetTempPath(), "ExcelMailMerge"),
        };

        foreach (var cand in candidates)
        {
            try
            {
                if (!Directory.Exists(cand)) Directory.CreateDirectory(cand);
                // 写入测试：确认目录确实可写
                var testFile = Path.Combine(cand, ".write_test");
                File.WriteAllText(testFile, "1");
                File.Delete(testFile);
                _settingsFolder = cand;
                _settingsFilePath = Path.Combine(cand, "settings.json");
                return;
            }
            catch
            {
                // 当前候选目录不可用，尝试下一个
            }
        }

        // 兜底：全部失败时使用临时目录路径（Load/Save 自身有 try-catch，不会导致应用崩溃）
        _settingsFolder = Path.Combine(Path.GetTempPath(), "ExcelMailMerge");
        _settingsFilePath = Path.Combine(_settingsFolder, "settings.json");
    }

    public AppSettings Load()
    {
        try
        {
            if (!File.Exists(_settingsFilePath)) return new AppSettings();
            var json = File.ReadAllText(_settingsFilePath);
            return JsonConvert.DeserializeObject<AppSettings>(json) ?? new AppSettings();
        }
        catch
        {
            return new AppSettings();
        }
    }

    public void Save(AppSettings settings)
    {
        try
        {
            var json = JsonConvert.SerializeObject(settings, Formatting.Indented);
            File.WriteAllText(_settingsFilePath, json);
        }
        catch
        {
            // 配置保存失败静默处理，不影响主流程
        }
    }
}
