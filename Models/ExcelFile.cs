using System.IO;

namespace ExcelMatcher.Models;

/// <summary>
///     Excel文件模型类，表示一个Excel文件及其基本信息
/// </summary>
public class ExcelFile
{
    /// <summary>
    ///     文件路径
    /// </summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>
    ///     文件密码（如果有加密）
    /// </summary>
    public string Password { get; set; } = string.Empty;

    /// <summary>
    ///     是否有密码加密
    /// </summary>
    public bool IsEncrypted => !string.IsNullOrEmpty(Password);

    /// <summary>
    ///     文件名（不含路径）
    /// </summary>
    public string FileName => Path.GetFileName(FilePath);

    /// <summary>
    ///     工作表列表
    /// </summary>
    public List<string> Worksheets { get; set; } = new();

    /// <summary>
    ///     当前选中的工作表
    /// </summary>
    public string SelectedWorksheet { get; set; } = string.Empty;

    /// <summary>
    ///     工作表的列名列表
    /// </summary>
    public List<string> Columns { get; set; } = new();

    /// <summary>
    ///     总行数
    /// </summary>
    public int RowCount { get; set; }

    /// <summary>
    ///     列数
    /// </summary>
    public int ColumnCount { get; set; }

    /// <summary>
    ///     工作表数量
    /// </summary>
    public int WorksheetCount => Worksheets.Count;

    /// <summary>
    ///     文件是否已加载
    /// </summary>
    public bool IsLoaded => !string.IsNullOrEmpty(FilePath) && Worksheets.Count > 0;

    /// <summary>
    ///     选中的多个工作表
    /// </summary>
    public List<string> SelectedWorksheets { get; set; } = new();

    /// <summary>
    ///     工作表信息字典 (工作表名 -> 行列信息)
    /// </summary>
    public Dictionary<string, (int RowCount, int ColumnCount, List<string> Columns)> WorksheetInfo { get; set; } =
        new();

    /// <summary>
    ///     最后检查文件的时间
    /// </summary>
    public DateTime LastChecked { get; set; } = DateTime.MinValue;
}