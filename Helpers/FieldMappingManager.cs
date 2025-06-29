using System.Data;
using ExcelMatcher.Models;

namespace ExcelMatcher.Helpers;

/// <summary>
///     字段映射管理器，处理字段映射和值转换的核心逻辑
/// </summary>
public class FieldMappingManager
{
    /// <summary>
    ///     验证字段映射是否有效
    /// </summary>
    /// <param name="mappings">字段映射列表</param>
    /// <param name="sourceColumns">源数据列</param>
    /// <param name="targetColumns">目标数据列</param>
    /// <returns>验证结果，如果无效则包含错误消息</returns>
    public static (bool IsValid, string ErrorMessage) ValidateFieldMappings(
        List<FieldMapping> mappings, IEnumerable<string> sourceColumns, IEnumerable<string> targetColumns)
    {
        if (mappings == null || !mappings.Any()) return (false, "没有定义字段映射");

        foreach (var mapping in mappings)
        {
            if (string.IsNullOrEmpty(mapping.SourceField)) return (false, "映射中的源字段不能为空");

            if (string.IsNullOrEmpty(mapping.TargetField)) return (false, "映射中的目标字段不能为空");

            if (!sourceColumns.Contains(mapping.SourceField)) return (false, $"源表中不存在字段 '{mapping.SourceField}'");
        }

        // 检查目标字段是否有重复（除非是向目标表添加新列）
        var duplicateTargets = mappings
            .Where(m => targetColumns.Contains(m.TargetField))
            .GroupBy(m => m.TargetField)
            .Where(g => g.Count() > 1)
            .Select(g => g.Key)
            .ToList();

        if (duplicateTargets.Any()) return (false, $"存在重复的目标字段映射: {string.Join(", ", duplicateTargets)}");

        return (true, string.Empty);
    }

    /// <summary>
    ///     应用字段映射到数据行
    /// </summary>
    /// <param name="sourceRow">源数据行</param>
    /// <param name="targetRow">目标数据行</param>
    /// <param name="mappings">字段映射列表</param>
    public static void ApplyMappings(DataRow sourceRow, DataRow targetRow, List<FieldMapping> mappings)
    {
        foreach (var mapping in mappings)
            try
            {
                // 确保目标行有对应的列
                if (!targetRow.Table.Columns.Contains(mapping.TargetField))
                    // 添加新列
                    targetRow.Table.Columns.Add(mapping.TargetField);

                // 复制值
                if (sourceRow.Table.Columns.Contains(mapping.SourceField))
                    targetRow[mapping.TargetField] = sourceRow[mapping.SourceField];
            }
            catch (Exception ex)
            {
                Logger.Error($"应用字段映射时出错: {mapping.SourceField} -> {mapping.TargetField}", ex);
                throw;
            }
    }

    /// <summary>
    ///     预览字段映射结果
    /// </summary>
    /// <param name="sourceTable">源数据表</param>
    /// <param name="targetTable">目标数据表</param>
    /// <param name="mappings">字段映射列表</param>
    /// <param name="maxRows">最大预览行数</param>
    /// <returns>包含映射结果的数据表</returns>
    public static DataTable PreviewMappings(DataTable sourceTable, DataTable targetTable, List<FieldMapping> mappings,
        int maxRows = 10)
    {
        // 创建预览表结构
        var previewTable = new DataTable();

        // 添加目标表所有列
        foreach (DataColumn column in targetTable.Columns) previewTable.Columns.Add(column.ColumnName, column.DataType);

        // 添加映射中可能新增的列
        foreach (var mapping in mappings)
            if (!previewTable.Columns.Contains(mapping.TargetField))
                previewTable.Columns.Add(mapping.TargetField);

        // 添加行数据（为了简单起见，只复制目标表的部分行）
        var rowCount = Math.Min(targetTable.Rows.Count, maxRows);
        for (var i = 0; i < rowCount; i++)
        {
            var newRow = previewTable.NewRow();

            // 复制原始目标表数据
            foreach (DataColumn column in targetTable.Columns)
                newRow[column.ColumnName] = targetTable.Rows[i][column.ColumnName];

            // 应用映射（假设第一行作为示例源）
            if (sourceTable.Rows.Count > 0)
                foreach (var mapping in mappings)
                    if (sourceTable.Columns.Contains(mapping.SourceField) &&
                        previewTable.Columns.Contains(mapping.TargetField))
                        // 为了演示，我们使用源表的第一行数据
                        newRow[mapping.TargetField] = sourceTable.Rows[0][mapping.SourceField];

            previewTable.Rows.Add(newRow);
        }

        return previewTable;
    }
}