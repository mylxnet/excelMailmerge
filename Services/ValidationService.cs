using ExcelMailMerge.Models;

namespace ExcelMailMerge.Services;

/// <summary>
/// 预校验服务：生成前集中检查所有问题，列出清单
/// </summary>
public class ValidationService
{
    private readonly NamingService _naming;

    public ValidationService(NamingService naming) { _naming = naming; }

    /// <summary>
    /// 数据源阶段校验（合并单元格检测、列名大括号、空数据等）
    /// </summary>
    public ValidationResult ValidateDataSource(ParsedDataSource dataSource)
    {
        var result = new ValidationResult();
        if (dataSource == null) return result;

        // 1. 列名含大括号 → 错误
        foreach (var col in dataSource.Columns)
        {
            if (col.DisplayName.Contains('{') || col.DisplayName.Contains('}'))
            {
                result.Issues.Add(new ValidationIssue
                {
                    Level = ValidationLevel.Error,
                    RelatedColumnName = col.OriginalName,
                    Message = $"列名 [{col.DisplayName}] 中包含非法的大括号字符 '{{' 或 '}}'",
                    Suggestion = "请修改数据源中的列名，去掉大括号字符"
                });
            }
        }

        // 2. 空标题列 → 警告
        foreach (var col in dataSource.Columns)
        {
            if (string.IsNullOrWhiteSpace(col.DisplayName))
            {
                result.Issues.Add(new ValidationIssue
                {
                    Level = ValidationLevel.Warning,
                    RelatedColumnName = col.OriginalName,
                    Message = "存在空标题列（标题行该单元格为空）",
                    Suggestion = "请在数据源标题行补全该列的列名，或调整标题行号"
                });
            }
        }

        // 3. 无数据行 → 错误
        if (dataSource.Rows.Count == 0)
        {
            result.Issues.Add(new ValidationIssue
            {
                Level = ValidationLevel.Error,
                Message = "未检测到任何数据行",
                Suggestion = "请确认标题行号是否正确，且数据行紧跟在标题行下方"
            });
        }

        return result;
    }

    public ValidationResult Validate(
        ParsedDataSource dataSource,
        ScannedTemplate template,
        List<string> namingColumns,
        string namingSeparator,
        OutputMode outputMode)
    {
        var result = new ValidationResult();

        // 1. 基础文件存在性（外部已保证，此处做兜底）
        if (dataSource == null || dataSource.Columns.Count == 0)
        {
            result.Issues.Add(new ValidationIssue
            {
                Level = ValidationLevel.Error,
                Message = "数据源未正确加载（无列名）",
                Suggestion = "请重新选择数据源文件并确认标题行号是否正确"
            });
            return result;
        }
        if (template == null || template.SheetCount == 0)
        {
            result.Issues.Add(new ValidationIssue
            {
                Level = ValidationLevel.Error,
                Message = "模板未正确加载（无Sheet）",
                Suggestion = "请重新选择模板文件"
            });
            return result;
        }
        if (dataSource.Rows.Count == 0)
        {
            result.Issues.Add(new ValidationIssue
            {
                Level = ValidationLevel.Error,
                Message = "数据源没有检测到任何数据行",
                Suggestion = "请确认标题行号是否选对，以及数据行是否紧跟在标题行下"
            });
            return result;
        }

        // 2. 模板Sheet>1 且 选了模式A → 警告（不报错）
        if (outputMode == OutputMode.SingleFileMultiSheet && template.SheetCount > 1)
        {
            result.Issues.Add(new ValidationIssue
            {
                Level = ValidationLevel.Warning,
                Message = $"模板共有{template.SheetCount}个Sheet，当前选择「所有行→同一文件」模式。" +
                          $"若有{dataSource.Rows.Count}行数据，则输出文件将产生 {template.SheetCount}×{dataSource.Rows.Count}={template.SheetCount * dataSource.Rows.Count} 个Sheet，" +
                          $"可能导致文件过大或Excel性能不佳。",
                Suggestion = "建议改用「每行→独立文件」模式，或确保数据量较小（<100行）"
            });
        }

        // 3. 占位符缺失列（警告，不报错）
        foreach (var ph in template.Placeholders.Where(p => !p.IsMatched))
        {
            result.Issues.Add(new ValidationIssue
            {
                Level = ValidationLevel.Warning,
                RelatedColumnName = ph.Key,
                Message = $"模板占位符 {{{ph.Key}}} 在数据源中找不到对应的列（出现{ph.OccurrenceCount}次），生成时该位置将被留空",
                Suggestion = "请检查模板占位符拼写是否与列名一致，或在数据源中新增对应列"
            });
        }

        // 4. 命名列校验
        if (namingColumns == null || namingColumns.Count == 0)
        {
            result.Issues.Add(new ValidationIssue
            {
                Level = ValidationLevel.Error,
                Message = "请至少选择一列作为「命名列」（用于生成文件名/Sheet名）",
                Suggestion = "在「命名规则」区域勾选1个或多个列，多列会自动组合"
            });
        }
        else
        {
            // 4.1 检查用户勾选的命名列是否都存在
            foreach (var nc in namingColumns)
            {
                if (!dataSource.DisplayNameSet.Contains(nc))
                {
                    result.Issues.Add(new ValidationIssue
                    {
                        Level = ValidationLevel.Error,
                        RelatedColumnName = nc,
                        Message = $"命名列 [{nc}] 在当前数据源中不存在（可能切换了Sheet或改了标题行号）",
                        Suggestion = "请重新勾选命名列"
                    });
                }
            }

            // 4.2 检查每行命名组合是否为空、是否重名
            if (result.ErrorCount == 0)
            {
                var (_, empties, dupGroups) = _naming.ValidateAllNames(dataSource, namingColumns, namingSeparator);

                foreach (var rowIdx in empties)
                {
                    result.Issues.Add(new ValidationIssue
                    {
                        Level = ValidationLevel.Error,
                        RelatedRowIndex = rowIdx,
                        Message = $"命名组合为空：该行中所有被选为命名列的值都是空的",
                        Suggestion = "请在该行的命名列中填入内容，或增加新的列参与命名组合"
                    });
                }

                foreach (var grp in dupGroups)
                {
                    result.Issues.Add(new ValidationIssue
                    {
                        Level = ValidationLevel.Error,
                        RelatedRowIndex = grp.First(),
                        Message = $"命名组合重复：名称「{grp.Key}」在以下行重复出现：第{string.Join("、第", grp)}行。" +
                                  $"系统不会自动加序号，必须保证命名唯一。",
                        Suggestion = "请修改上述行中命名列的数据，或点击「+ 添加列参与命名」增加新列（如身份证号、流水号）以区分"
                    });
                }
            }
        }

        // 5. 检查列名中是否有大括号（DataSourceService.Parse 时会抛异常，这里兜底做信息级提示）
        foreach (var col in dataSource.Columns)
        {
            if (col.DisplayName.Contains('{') || col.DisplayName.Contains('}'))
            {
                result.Issues.Add(new ValidationIssue
                {
                    Level = ValidationLevel.Error,
                    RelatedColumnName = col.OriginalName,
                    Message = $"列名 [{col.DisplayName}] 中包含非法的大括号字符 '{{' 或 '}}'",
                    Suggestion = "请修改数据源中的列名，去掉大括号字符"
                });
            }
        }

        return result;
    }
}
