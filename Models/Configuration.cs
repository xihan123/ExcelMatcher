using Newtonsoft.Json;

namespace ExcelMatcher.Models;

/// <summary>
///     配置模型类，用于保存和加载用户配置
/// </summary>
public class Configuration
{
    /// <summary>
    ///     配置名称
    /// </summary>
    public string Name { get; set; } = DateTime.Now.ToString("yyyyMMdd_HHmmss");

    /// <summary>
    ///     配置文件路径
    /// </summary>
    [JsonIgnore]
    public string Path { get; set; } = string.Empty;

    /// <summary>
    ///     主表文件路径
    /// </summary>
    public string PrimaryFilePath { get; set; } = string.Empty;

    /// <summary>
    ///     主表文件密码
    /// </summary>
    public string PrimaryFilePassword { get; set; } = string.Empty;

    /// <summary>
    ///     主表工作表（单工作表兼容）
    /// </summary>
    public string PrimaryWorksheet { get; set; } = string.Empty;

    /// <summary>
    ///     主表多工作表集合
    /// </summary>
    public List<string> PrimaryWorksheets { get; set; } = new();

    /// <summary>
    ///     辅助表文件路径
    /// </summary>
    public string SecondaryFilePath { get; set; } = string.Empty;

    /// <summary>
    ///     辅助表文件密码
    /// </summary>
    public string SecondaryFilePassword { get; set; } = string.Empty;

    /// <summary>
    ///     辅助表工作表（单工作表兼容）
    /// </summary>
    public string SecondaryWorksheet { get; set; } = string.Empty;

    /// <summary>
    ///     辅助表多工作表集合
    /// </summary>
    public List<string> SecondaryWorksheets { get; set; } = new();

    /// <summary>
    ///     主表匹配字段
    /// </summary>
    public List<string> PrimaryMatchFields { get; set; } = new();

    /// <summary>
    ///     辅助表匹配字段
    /// </summary>
    public List<string> SecondaryMatchFields { get; set; } = new();

    /// <summary>
    ///     字段映射列表
    /// </summary>
    public List<FieldMapping> FieldMappings { get; set; } = new();

    /// <summary>
    ///     主表筛选条件
    /// </summary>
    public List<FilterCondition> PrimaryFilters { get; set; } = new();

    /// <summary>
    ///     辅助表筛选条件
    /// </summary>
    public List<FilterCondition> SecondaryFilters { get; set; } = new();
}