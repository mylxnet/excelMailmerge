using ExcelMailMerge.Services;

namespace ExcelMailMerge;

/// <summary>
/// 简易服务定位器（不引入第三方DI容器，保持轻量）
/// </summary>
public static class ServiceLocator
{
    public static SettingsService Settings { get; } = new();
    public static DataSourceService DataSource { get; } = new();
    public static TemplateService Template { get; } = new();
    public static NamingService Naming { get; } = new();
    public static ValidationService Validation { get; } = new(Naming);
    public static GenerationEngine Engine { get; } = new(Naming, Template);
}
