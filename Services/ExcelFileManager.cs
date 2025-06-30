using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text;
using ExcelMatcher.Models;
using OfficeOpenXml;

namespace ExcelMatcher.Services;

/// <summary>
///     Excel文件管理服务，负责Excel文件的读写操作
/// </summary>
public class ExcelFileManager
{
    // 设置EPPlus许可
    static ExcelFileManager()
    {
        // 设置EPPlus的LicenseContext为非商业用途
        ExcelPackage.License.SetNonCommercialOrganization("xihan123");
    }

    /// <summary>
    ///     加载Excel文件
    /// </summary>
    public async Task<ExcelFile> LoadExcelFileAsync(string filePath, string? password = null)
    {
        if (string.IsNullOrEmpty(filePath))
            throw new ArgumentNullException(nameof(filePath));

        if (!File.Exists(filePath))
            throw new FileNotFoundException($"文件不存在: {filePath}");

        return await Task.Run(() =>
        {
            try
            {
                // 刷新文件信息以获取最新的修改时间
                var fileInfo = new FileInfo(filePath);
                fileInfo.Refresh();

                var excelFile = new ExcelFile
                {
                    FilePath = filePath,
                    Password = password,
                    // 更新最后检查时间
                    LastChecked = fileInfo.LastWriteTime
                };

                using (var package = GetExcelPackage(filePath, password))
                {
                    // 获取所有工作表信息
                    foreach (var worksheet in package.Workbook.Worksheets) excelFile.Worksheets.Add(worksheet.Name);
                }

                LogDebug($"已加载Excel文件: {filePath}, 最后修改时间: {excelFile.LastChecked}, 工作表数量: {excelFile.WorksheetCount}");
                return excelFile;
            }
            catch (Exception ex)
            {
                LogDebug($"加载Excel文件时出错: {filePath}, 错误: {ex.Message}");
                throw;
            }
        });
    }

    /// <summary>
    ///     记录调试信息
    /// </summary>
    private void LogDebug(string message)
    {
        Debug.WriteLine($"{DateTime.Now:HH:mm:ss.fff} - {message}");
    }

    /// <summary>
    ///     彻底关闭Excel文件
    /// </summary>
    public Task CloseFileAsync(ExcelFile file)
    {
        if (file == null)
            return Task.CompletedTask;

        return Task.Run(() =>
        {
            try
            {
                // 清理内存中的数据
                file.Worksheets.Clear();
                file.Columns.Clear();
                file.SelectedWorksheets.Clear();
                file.WorksheetInfo.Clear();
                file.RowCount = 0;
                file.ColumnCount = 0;

                // 尝试额外的资源清理
                for (var i = 0; i < 3; i++)
                {
                    // 强制进行垃圾回收
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    // 尝试检测文件是否被锁定
                    try
                    {
                        if (!string.IsNullOrEmpty(file.FilePath) && File.Exists(file.FilePath))
                            using (var fs = File.Open(file.FilePath, FileMode.Open, FileAccess.Read,
                                       FileShare.ReadWrite))
                            {
                                // 如果能打开文件，说明资源已释放
                            }

                        break;
                    }
                    catch
                    {
                        // 文件可能仍被锁定，等待一会再尝试
                        Thread.Sleep(100);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"关闭Excel文件时出错: {ex.Message}");
            }
        });
    }

    /// <summary>
    ///     加载工作表信息
    /// </summary>
    /// <param name="excelFile">Excel文件对象</param>
    /// <returns>更新后的Excel文件对象</returns>
    public async Task<ExcelFile> LoadWorksheetInfoAsync(ExcelFile excelFile)
    {
        if (string.IsNullOrEmpty(excelFile.SelectedWorksheet))
            throw new ArgumentException("未选择工作表");

        try
        {
            await Task.Run(() =>
            {
                using (var package = GetExcelPackage(excelFile.FilePath, excelFile.Password))
                {
                    var worksheet = package.Workbook.Worksheets[excelFile.SelectedWorksheet];
                    if (worksheet == null)
                        throw new ArgumentException($"找不到名为 {excelFile.SelectedWorksheet} 的工作表");

                    // 清空现有列名
                    excelFile.Columns.Clear();

                    // 获取表头（假设第一行是表头）
                    var colCount = worksheet.Dimension?.End.Column ?? 0;
                    for (var col = 1; col <= colCount; col++)
                    {
                        var headerValue = worksheet.Cells[1, col].Text;
                        if (!string.IsNullOrEmpty(headerValue)) excelFile.Columns.Add(headerValue);
                    }

                    // 获取行数（减去表头）
                    excelFile.RowCount = (worksheet.Dimension?.End.Row ?? 1) - 1;
                    excelFile.ColumnCount = excelFile.Columns.Count;
                }
            });

            return excelFile;
        }
        catch (Exception ex)
        {
            throw new Exception($"加载工作表 {excelFile.SelectedWorksheet} 信息时出错: {ex.Message}", ex);
        }
    }

    /// <summary>
    ///     获取工作表数据
    /// </summary>
    /// <param name="excelFile">Excel文件对象</param>
    /// <param name="maxRows">最大行数（0表示所有行）</param>
    /// <returns>数据表</returns>
    public async Task<DataTable> GetWorksheetDataAsync(ExcelFile excelFile, int maxRows = 100)
    {
        if (string.IsNullOrEmpty(excelFile.SelectedWorksheet))
            throw new ArgumentException("未选择工作表");

        var dataTable = new DataTable();

        try
        {
            await Task.Run(() =>
            {
                using (var package = GetExcelPackage(excelFile.FilePath, excelFile.Password))
                {
                    var worksheet = package.Workbook.Worksheets[excelFile.SelectedWorksheet];
                    if (worksheet == null)
                        throw new ArgumentException($"找不到名为 {excelFile.SelectedWorksheet} 的工作表");

                    var startRow = 1;
                    var startCol = 1;
                    var endRow = worksheet.Dimension?.End.Row ?? 0;
                    var endCol = worksheet.Dimension?.End.Column ?? 0;

                    // 限制最大行数
                    if (maxRows > 0 && endRow > maxRows + startRow) endRow = maxRows + startRow;

                    // 添加列
                    for (var col = startCol; col <= endCol; col++)
                    {
                        var columnName = worksheet.Cells[startRow, col].Text;
                        if (string.IsNullOrEmpty(columnName))
                            columnName = $"列{col}";

                        dataTable.Columns.Add(columnName);
                    }

                    // 添加数据行
                    for (var row = startRow + 1; row <= endRow; row++)
                    {
                        var dataRow = dataTable.NewRow();
                        for (var col = startCol; col <= endCol; col++)
                            dataRow[col - startCol] = worksheet.Cells[row, col].Text;
                        dataTable.Rows.Add(dataRow);
                    }
                }
            });

            return dataTable;
        }
        catch (Exception ex)
        {
            throw new Exception($"获取工作表 {excelFile.SelectedWorksheet} 数据时出错: {ex.Message}", ex);
        }
    }

    /// <summary>
    ///     执行数据合并操作
    /// </summary>
    /// <param name="primaryFile">主表文件</param>
    /// <param name="secondaryFile">辅助表文件</param>
    /// <param name="primaryMatchFields">主表匹配字段</param>
    /// <param name="secondaryMatchFields">辅助表匹配字段</param>
    /// <param name="fieldMappings">字段映射关系</param>
    /// <param name="primaryFilters">主表筛选条件</param>
    /// <param name="secondaryFilters">辅助表筛选条件</param>
    /// <param name="progressCallback">进度回调</param>
    /// <returns>合并结果统计</returns>
    public async Task<(int ProcessedRows, int MatchedRows, int NewColumnsAdded, string OutputPath)>
        MergeExcelFilesAsync(
            ExcelFile primaryFile,
            ExcelFile secondaryFile,
            List<string> primaryMatchFields,
            List<string> secondaryMatchFields,
            List<FieldMapping> fieldMappings,
            List<FilterCondition> primaryFilters,
            List<FilterCondition> secondaryFilters,
            IProgress<(int Current, int Total, string Message)> progressCallback)
    {
        if (primaryFile == null || secondaryFile == null)
            throw new ArgumentNullException("Excel文件对象不能为空");

        if (string.IsNullOrEmpty(primaryFile.SelectedWorksheet) ||
            string.IsNullOrEmpty(secondaryFile.SelectedWorksheet))
            throw new ArgumentException("未选择工作表");

        if (primaryMatchFields == null || primaryMatchFields.Count == 0 ||
            secondaryMatchFields == null || secondaryMatchFields.Count == 0 ||
            primaryMatchFields.Count != secondaryMatchFields.Count)
            throw new ArgumentException("匹配字段不能为空且数量必须相同");

        if (fieldMappings == null || fieldMappings.Count == 0)
            throw new ArgumentException("字段映射不能为空");

        var processedRows = 0;
        var matchedRows = 0;
        var newColumnsAdded = 0;
        var outputPath = string.Empty;

        try
        {
            // 创建新文件
            var directory = Path.GetDirectoryName(primaryFile.FilePath) ?? "";
            var fileName = Path.GetFileNameWithoutExtension(primaryFile.FilePath);
            var extension = Path.GetExtension(primaryFile.FilePath);
            outputPath = Path.Combine(directory, $"{fileName}_合并结果_{DateTime.Now:yyyyMMddHHmmss}{extension}");

            // 复制主文件作为结果文件
            File.Copy(primaryFile.FilePath, outputPath, true);

            // 加载主表数据
            progressCallback?.Report((10, 100, "正在加载主表数据..."));
            var primaryData = await GetWorksheetDataAsync(primaryFile, 0);
            primaryData = ApplyFilters(primaryData, primaryFilters);

            // 加载辅助表数据
            progressCallback?.Report((30, 100, "正在加载辅助表数据..."));
            var secondaryData = await GetWorksheetDataAsync(secondaryFile, 0);
            secondaryData = ApplyFilters(secondaryData, secondaryFilters);

            // 收集主表和辅助表所有列
            var primaryColumns = new HashSet<string>(primaryData.Columns.Cast<DataColumn>().Select(c => c.ColumnName));
            var secondaryColumns =
                new HashSet<string>(secondaryData.Columns.Cast<DataColumn>().Select(c => c.ColumnName));

            // 验证字段映射
            var validMappings = ValidateFieldMappings(fieldMappings, secondaryColumns, primaryColumns);

            // 建立辅助表索引，用于查找匹配记录
            progressCallback?.Report((40, 100, "正在建立辅助表索引..."));
            var secondaryIndex = new Dictionary<string, DataRow>();
            foreach (DataRow row in secondaryData.Rows)
            {
                var key = GetCompositeKey(row, secondaryMatchFields);
                if (!string.IsNullOrEmpty(key) && !secondaryIndex.ContainsKey(key)) secondaryIndex.Add(key, row);
            }

            // 匹配和合并数据
            progressCallback?.Report((50, 100, "正在匹配和合并数据..."));
            var totalRows = primaryData.Rows.Count;
            for (var i = 0; i < totalRows; i++)
            {
                var primaryRow = primaryData.Rows[i];
                var key = GetCompositeKey(primaryRow, primaryMatchFields);

                if (!string.IsNullOrEmpty(key) && secondaryIndex.TryGetValue(key, out var secondaryRow))
                {
                    // 复制字段值
                    foreach (var mapping in validMappings)
                        try
                        {
                            // 验证源字段存在
                            if (!secondaryData.Columns.Contains(mapping.SourceField)) continue;

                            // 准备目标字段
                            if (!primaryData.Columns.Contains(mapping.TargetField))
                                primaryData.Columns.Add(mapping.TargetField);

                            // 复制数据
                            primaryRow[mapping.TargetField] = secondaryRow[mapping.SourceField];
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine(
                                $"应用字段映射时出错: {mapping.SourceField} -> {mapping.TargetField}, 原因: {ex.Message}");
                        }

                    matchedRows++;
                }

                processedRows++;

                if (i % 100 == 0 || i == totalRows - 1)
                {
                    var progress = 50 + (i + 1) * 30 / totalRows;
                    progressCallback?.Report((progress, 100, $"正在处理数据... {i + 1}/{totalRows}"));
                }
            }

            // 保存结果
            progressCallback?.Report((80, 100, "正在保存结果..."));
            using (var package = GetExcelPackage(outputPath, primaryFile.Password))
            {
                var worksheet = package.Workbook.Worksheets[primaryFile.SelectedWorksheet];
                if (worksheet == null)
                    throw new InvalidOperationException($"工作表 {primaryFile.SelectedWorksheet} 不存在");

                // 检查并添加新列
                var headerRow = 1; // 假设第一行是表头
                var existingColumns = new Dictionary<string, int>();
                var lastColumn = worksheet.Dimension?.End.Column ?? 0;

                // 读取现有列名和列索引
                for (var col = 1; col <= lastColumn; col++)
                {
                    var headerText = worksheet.Cells[headerRow, col].Text;
                    if (!string.IsNullOrEmpty(headerText)) existingColumns[headerText] = col;
                }

                // 添加新列
                foreach (var mapping in validMappings)
                    if (!existingColumns.ContainsKey(mapping.TargetField))
                    {
                        lastColumn++;
                        worksheet.Cells[headerRow, lastColumn].Value = mapping.TargetField;
                        existingColumns[mapping.TargetField] = lastColumn;
                        newColumnsAdded++;
                    }

                // 清除工作表现有数据但保留表头
                var startRow = 2; // 跳过表头行
                var endRow = worksheet.Dimension?.End.Row ?? 0;
                if (endRow >= startRow) worksheet.Cells[startRow, 1, endRow, worksheet.Dimension.End.Column].Clear();

                // 写入新数据到工作表
                for (var i = 0; i < primaryData.Rows.Count; i++)
                {
                    var dataRow = primaryData.Rows[i];
                    for (var j = 0; j < primaryData.Columns.Count; j++)
                        try
                        {
                            var columnName = primaryData.Columns[j].ColumnName;
                            if (existingColumns.TryGetValue(columnName, out var colIndex))
                                worksheet.Cells[i + startRow, colIndex].Value = dataRow[j];
                        }
                        catch (Exception ex)
                        {
                            // 记录错误但继续处理其他单元格
                            Debug.WriteLine($"写入单元格数据时出错: 行={i}, 列={j}, 原因: {ex.Message}");
                        }
                }

                progressCallback?.Report((90, 100, "正在保存文件..."));
                package.Save();
            }

            progressCallback?.Report((100, 100, "完成"));
            return (processedRows, matchedRows, newColumnsAdded, outputPath);
        }
        catch (Exception ex)
        {
            throw new Exception($"合并Excel文件时出错: {ex.Message}", ex);
        }
    }

    /// <summary>
    ///     检查字段映射的有效性
    /// </summary>
    private List<FieldMapping> ValidateFieldMappings(List<FieldMapping> mappings, HashSet<string> sourceColumns,
        HashSet<string> targetColumns)
    {
        var validMappings = new List<FieldMapping>();

        foreach (var mapping in mappings)
        {
            // 检查源字段是否存在
            if (!sourceColumns.Contains(mapping.SourceField))
            {
                Debug.WriteLine($"警告: 源字段 '{mapping.SourceField}' 在辅助表中不存在，已跳过此映射");
                continue;
            }

            // 对于目标字段，如果不存在则会创建新列，所以不需要验证存在性
            validMappings.Add(mapping);
        }

        return validMappings;
    }

    /// <summary>
    ///     从数据行获取复合键
    /// </summary>
    private string GetCompositeKey(DataRow row, List<string> fields)
    {
        if (row == null || fields == null || fields.Count == 0)
            return string.Empty;

        var keyParts = new List<string>();
        foreach (var field in fields)
            if (row.Table.Columns.Contains(field) && row[field] != null && row[field] != DBNull.Value)
                keyParts.Add(row[field].ToString().Trim());
            else
                keyParts.Add(string.Empty);

        return string.Join("||", keyParts);
    }

    /// <summary>
    ///     应用筛选条件
    /// </summary>
    public DataTable ApplyFilters(DataTable dataTable, List<FilterCondition> filters)
    {
        if (filters == null || filters.Count == 0)
            return dataTable;

        var result = dataTable.Copy();
        var filterExpression = BuildFilterExpression(filters, dataTable);

        if (string.IsNullOrEmpty(filterExpression))
            return result;

        try
        {
            var rows = dataTable.Select(filterExpression);
            result.Clear();
            foreach (var row in rows) result.ImportRow(row);
        }
        catch (Exception ex)
        {
            throw new Exception($"应用筛选条件时出错: {ex.Message}", ex);
        }

        return result;
    }

    /// <summary>
    ///     构建筛选表达式
    /// </summary>
    private string BuildFilterExpression(List<FilterCondition> filters, DataTable dataTable)
    {
        if (filters == null || filters.Count == 0)
            return string.Empty;

        var expressions = new List<string>();

        foreach (var filter in filters)
        {
            if (!dataTable.Columns.Contains(filter.Field))
                continue;

            var expression = string.Empty;
            switch (filter.Operator)
            {
                case FilterOperator.Equals:
                    expression = $"[{filter.Field}] = '{filter.Value.Replace("'", "''")}'";
                    break;
                case FilterOperator.NotEquals:
                    expression = $"[{filter.Field}] <> '{filter.Value.Replace("'", "''")}'";
                    break;
                case FilterOperator.Contains:
                    expression = $"[{filter.Field}] LIKE '%{filter.Value.Replace("'", "''")}%'";
                    break;
                case FilterOperator.NotContains:
                    expression = $"[{filter.Field}] NOT LIKE '%{filter.Value.Replace("'", "''")}%'";
                    break;
                case FilterOperator.StartsWith:
                    expression = $"[{filter.Field}] LIKE '{filter.Value.Replace("'", "''")}%'";
                    break;
                case FilterOperator.EndsWith:
                    expression = $"[{filter.Field}] LIKE '%{filter.Value.Replace("'", "''")}'";
                    break;
                case FilterOperator.GreaterThan:
                    expression = $"[{filter.Field}] > '{filter.Value.Replace("'", "''")}'";
                    break;
                case FilterOperator.LessThan:
                    expression = $"[{filter.Field}] < '{filter.Value.Replace("'", "''")}'";
                    break;
                case FilterOperator.GreaterThanOrEqual:
                    expression = $"[{filter.Field}] >= '{filter.Value.Replace("'", "''")}'";
                    break;
                case FilterOperator.LessThanOrEqual:
                    expression = $"[{filter.Field}] <= '{filter.Value.Replace("'", "''")}'";
                    break;
                case FilterOperator.IsNull:
                    expression = $"[{filter.Field}] IS NULL OR [{filter.Field}] = ''";
                    break;
                case FilterOperator.IsNotNull:
                    expression = $"[{filter.Field}] IS NOT NULL AND [{filter.Field}] <> ''";
                    break;
            }

            if (!string.IsNullOrEmpty(expression)) expressions.Add(expression);
        }

        // 组合表达式
        var result = string.Empty;
        for (var i = 0; i < expressions.Count; i++)
        {
            if (i > 0)
            {
                var logicalOperator = filters[i].LogicalOperator == LogicalOperator.And ? " AND " : " OR ";
                result += logicalOperator;
            }

            result += expressions[i];
        }

        return result;
    }


    /// <summary>
    ///     获取Excel包对象
    /// </summary>
    private ExcelPackage GetExcelPackage(string filePath, string password)
    {
        if (string.IsNullOrEmpty(password)) return new ExcelPackage(new FileInfo(filePath));

        return new ExcelPackage(new FileInfo(filePath), password);
    }

    /// <summary>
    ///     执行多工作表数据合并操作
    /// </summary>
    public async Task<(int ProcessedRows, int MatchedRows, int NewColumnsAdded, string OutputPath)>
        MergeMultipleWorksheetsAsync(
            ExcelFile primaryFile,
            ExcelFile secondaryFile,
            List<string> primaryMatchFields,
            List<string> secondaryMatchFields,
            List<FieldMapping> fieldMappings,
            List<FilterCondition> primaryFilters,
            List<FilterCondition> secondaryFilters,
            IProgress<(int Current, int Total, string Message)> progressCallback)
    {
        if (primaryFile == null || secondaryFile == null)
            throw new ArgumentNullException("Excel文件对象不能为空");

        if (primaryFile.SelectedWorksheets.Count == 0 || secondaryFile.SelectedWorksheets.Count == 0)
            throw new ArgumentException("未选择工作表");

        var processedRows = 0;
        var matchedRows = 0;
        var newColumnsAdded = 0;

        var outputPath = string.Empty;
        // 创建新文件
        var directory = Path.GetDirectoryName(primaryFile.FilePath) ?? "";
        var fileName = Path.GetFileNameWithoutExtension(primaryFile.FilePath);
        var extension = Path.GetExtension(primaryFile.FilePath);
        outputPath = Path.Combine(directory, $"{fileName}_合并结果_{DateTime.Now:yyyyMMddHHmmss}{extension}");

        // 复制主文件作为结果文件
        File.Copy(primaryFile.FilePath, outputPath, true);

        try
        {
            // 复制主文件作为结果文件
            File.Copy(primaryFile.FilePath, outputPath, true);

            // 加载辅助表数据到内存
            progressCallback?.Report((10, 100, "正在加载辅助表数据..."));

            // 创建辅助表数据缓存
            var secondaryDataCache = new Dictionary<string, DataTable>();
            foreach (var worksheet in secondaryFile.SelectedWorksheets)
            {
                secondaryFile.SelectedWorksheet = worksheet;
                var data = await GetWorksheetDataAsync(secondaryFile, 0);
                data = ApplyFilters(data, secondaryFilters);
                secondaryDataCache[worksheet] = data;
            }

            // 创建辅助表索引（使用复合键）
            var secondaryIndex = new Dictionary<string, (string Worksheet, DataRow Row)>();
            progressCallback?.Report((20, 100, "正在建立辅助表索引..."));

            foreach (var entry in secondaryDataCache)
            foreach (DataRow row in entry.Value.Rows)
            {
                var key = GetCompositeKey(row, secondaryMatchFields);
                if (!string.IsNullOrEmpty(key) && !secondaryIndex.ContainsKey(key))
                    secondaryIndex.Add(key, (entry.Key, row));
            }

            // 处理每个主表工作表
            var currentProgress = 20;
            var progressPerWorksheet = 70 / primaryFile.SelectedWorksheets.Count;

            using (var package = GetExcelPackage(outputPath, primaryFile.Password))
            {
                foreach (var worksheetName in primaryFile.SelectedWorksheets)
                    try // 添加内部try-catch以单独处理每个工作表的错误
                    {
                        progressCallback?.Report((currentProgress, 100, $"正在处理工作表 {worksheetName}..."));

                        // 加载主表数据
                        primaryFile.SelectedWorksheet = worksheetName;
                        var primaryData = await GetWorksheetDataAsync(primaryFile, 0);
                        primaryData = ApplyFilters(primaryData, primaryFilters);

                        // 获取工作表引用
                        var worksheet = package.Workbook.Worksheets[worksheetName];
                        if (worksheet == null)
                        {
                            progressCallback?.Report((currentProgress, 100, $"跳过不存在的工作表 {worksheetName}"));
                            continue;
                        }

                        // 检查并添加新列
                        var headerRow = 1; // 假设第一行是表头
                        var existingColumns = new Dictionary<string, int>();
                        var lastColumn = worksheet.Dimension?.End.Column ?? 0;

                        // 读取现有列名和列索引
                        for (var col = 1; col <= lastColumn; col++)
                        {
                            var headerText = worksheet.Cells[headerRow, col].Text;
                            if (!string.IsNullOrEmpty(headerText)) existingColumns[headerText] = col;
                        }

                        // 添加新列
                        var newColumnsInThisSheet = 0;
                        // 1. 首先收集所有需要添加的映射字段
                        var mappingsToAdd = new Dictionary<string, string>();
                        foreach (var mapping in fieldMappings)
                            // 收集需添加的目标字段
                            if (!existingColumns.ContainsKey(mapping.TargetField))
                                mappingsToAdd[mapping.TargetField] = mapping.SourceField;

                        // 2. 将收集的字段添加为Excel表格的列
                        foreach (var targetField in mappingsToAdd.Keys)
                        {
                            lastColumn++;
                            worksheet.Cells[headerRow, lastColumn].Value = targetField;
                            existingColumns[targetField] = lastColumn;
                            newColumnsInThisSheet++;
                        }

                        newColumnsAdded += newColumnsInThisSheet;

                        // 执行数据匹配与合并
                        var totalRows = primaryData.Rows.Count;
                        var sheetMatchedRows = 0;

                        // 收集次表中存在的列字段
                        var secondaryTableColumns = new Dictionary<string, HashSet<string>>();
                        foreach (var worksheetPair in secondaryFile.WorksheetInfo)
                        {
                            var columns = new HashSet<string>();
                            foreach (var column in worksheetPair.Value.Columns) columns.Add(column);
                            secondaryTableColumns[worksheetPair.Key] = columns;
                        }

                        for (var i = 0; i < totalRows; i++)
                        {
                            var primaryRow = primaryData.Rows[i];
                            var key = GetCompositeKey(primaryRow, primaryMatchFields);

                            if (!string.IsNullOrEmpty(key) && secondaryIndex.TryGetValue(key, out var secondaryEntry))
                            {
                                // 复制字段值
                                var secondaryRow = secondaryEntry.Row;
                                var secondaryWorksheetName = secondaryEntry.Worksheet;

                                // 确保该辅助表工作表的列集合存在
                                if (!secondaryTableColumns.ContainsKey(secondaryWorksheetName)) continue;

                                var secondaryColumns = secondaryTableColumns[secondaryWorksheetName];

                                foreach (var mapping in fieldMappings)
                                {
                                    // 严格检查源列是否存在于辅助表
                                    if (!secondaryColumns.Contains(mapping.SourceField)) continue; // 跳过不存在的列

                                    // 检查源字段是否确实存在于此行所在的DataTable
                                    if (!secondaryRow.Table.Columns.Contains(mapping.SourceField)) continue; // 跳过不存在的列

                                    try
                                    {
                                        // 确保目标字段在DataTable中存在
                                        if (!primaryData.Columns.Contains(mapping.TargetField))
                                            primaryData.Columns.Add(mapping.TargetField);

                                        // 复制数值
                                        primaryRow[mapping.TargetField] = secondaryRow[mapping.SourceField];
                                    }
                                    catch (Exception ex)
                                    {
                                        // 记录错误但继续处理其他字段
                                        Debug.WriteLine(
                                            $"映射字段时出错: {mapping.SourceField} -> {mapping.TargetField}, 原因: {ex.Message}");
                                    }
                                }

                                sheetMatchedRows++;
                            }

                            processedRows++;

                            // 报告进度
                            if (i % 100 == 0 || i == totalRows - 1)
                            {
                                var progress = currentProgress + (i + 1) * progressPerWorksheet / totalRows;
                                progressCallback?.Report((progress, 100,
                                    $"正在处理工作表 {worksheetName} 数据... {i + 1}/{totalRows}"));
                            }
                        }

                        matchedRows += sheetMatchedRows;

                        // 清除工作表现有数据但保留表头
                        var startRow = 2; // 跳过表头行
                        var endRow = worksheet.Dimension?.End.Row ?? 0;
                        if (endRow >= startRow)
                            worksheet.Cells[startRow, 1, endRow, worksheet.Dimension.End.Column].Clear();

                        // 写入新数据到工作表
                        for (var i = 0; i < primaryData.Rows.Count; i++)
                        {
                            var dataRow = primaryData.Rows[i];
                            for (var j = 0; j < primaryData.Columns.Count; j++)
                                try
                                {
                                    var columnName = primaryData.Columns[j].ColumnName;
                                    if (existingColumns.TryGetValue(columnName, out var colIndex))
                                        worksheet.Cells[i + startRow, colIndex].Value = dataRow[j];
                                }
                                catch (Exception ex)
                                {
                                    // 记录错误但继续处理其他单元格
                                    Debug.WriteLine($"写入单元格数据时出错: 行={i}, 列={j}, 原因: {ex.Message}");
                                }
                        }

                        currentProgress += progressPerWorksheet;
                    }
                    catch (Exception ex)
                    {
                        // 记录单个工作表处理错误，但继续处理其他工作表
                        Debug.WriteLine($"处理工作表 {worksheetName} 时出错: {ex.Message}");
                        progressCallback?.Report((currentProgress, 100, $"处理工作表 {worksheetName} 时出错"));
                        currentProgress += progressPerWorksheet;
                    }

                // 保存结果
                progressCallback?.Report((90, 100, "正在保存结果..."));
                package.Save();
            }

            progressCallback?.Report((100, 100, "完成"));
            return (processedRows, matchedRows, newColumnsAdded, outputPath);
        }
        catch (Exception ex)
        {
            throw new Exception($"合并Excel文件时出错: {ex.Message}", ex);
        }
    }


    /// <summary>
    ///     收集所有工作表中的列名
    /// </summary>
    private HashSet<string> CollectAllColumns(
        Dictionary<string, (int RowCount, int ColumnCount, List<string> Columns)> worksheetInfo)
    {
        var allColumns = new HashSet<string>();

        foreach (var (_, _, columns) in worksheetInfo.Values)
        foreach (var column in columns)
            allColumns.Add(column);

        return allColumns;
    }

    #region 诊断工具

    /// <summary>
    ///     诊断匹配字段
    /// </summary>
    public async Task<string> DiagnoseMatchFieldsAsync(
        ExcelFile primaryFile,
        ExcelFile secondaryFile,
        List<string> primaryMatchFields,
        List<string> secondaryMatchFields)
    {
        try
        {
            var sb = new StringBuilder();
            sb.AppendLine("匹配字段诊断报告:");
            sb.AppendLine("===================");

            // 获取主表和辅助表数据
            var primaryData = await GetWorksheetDataAsync(primaryFile, 5); // 只取前5行进行诊断
            var secondaryData = await GetWorksheetDataAsync(secondaryFile, 5); // 只取前5行进行诊断

            sb.AppendLine($"主表字段: {string.Join(", ", primaryMatchFields)}");
            sb.AppendLine($"辅助表字段: {string.Join(", ", secondaryMatchFields)}");
            sb.AppendLine();

            // 检查主表数据
            sb.AppendLine("主表数据样本:");
            foreach (DataRow row in primaryData.Rows)
            {
                var key = GetCompositeKey(row, primaryMatchFields);
                sb.Append($"键: {key} - 值: [");
                foreach (var field in primaryMatchFields)
                    if (row.Table.Columns.Contains(field))
                    {
                        var val = row[field];
                        sb.Append($"{field}='{val}' (类型:{val?.GetType().Name ?? "null"}), ");
                    }
                    else
                    {
                        sb.Append($"{field}=不存在, ");
                    }

                sb.AppendLine("]");
            }

            sb.AppendLine();

            // 检查辅助表数据
            sb.AppendLine("辅助表数据样本:");
            foreach (DataRow row in secondaryData.Rows)
            {
                var key = GetCompositeKey(row, secondaryMatchFields);
                sb.Append($"键: {key} - 值: [");
                foreach (var field in secondaryMatchFields)
                    if (row.Table.Columns.Contains(field))
                    {
                        var val = row[field];
                        sb.Append($"{field}='{val}' (类型:{val?.GetType().Name ?? "null"}), ");
                    }
                    else
                    {
                        sb.Append($"{field}=不存在, ");
                    }

                sb.AppendLine("]");
            }

            // 检查潜在的匹配问题
            sb.AppendLine();
            sb.AppendLine("潜在问题检测:");

            // 检查字段类型是否匹配
            var primaryTypes = new Dictionary<string, Type>();
            var secondaryTypes = new Dictionary<string, Type>();

            for (var i = 0; i < primaryMatchFields.Count && i < secondaryMatchFields.Count; i++)
                if (primaryData.Columns.Contains(primaryMatchFields[i]) &&
                    secondaryData.Columns.Contains(secondaryMatchFields[i]))
                {
                    primaryTypes[primaryMatchFields[i]] = primaryData.Columns[primaryMatchFields[i]].DataType;
                    secondaryTypes[secondaryMatchFields[i]] = secondaryData.Columns[secondaryMatchFields[i]].DataType;

                    if (primaryTypes[primaryMatchFields[i]] != secondaryTypes[secondaryMatchFields[i]])
                        sb.AppendLine(
                            $"警告: 字段类型不匹配 - 主表 '{primaryMatchFields[i]}' ({primaryTypes[primaryMatchFields[i]].Name}) vs 辅助表 '{secondaryMatchFields[i]}' ({secondaryTypes[secondaryMatchFields[i]].Name})");
                }

            // 尝试匹配几个样本记录
            sb.AppendLine();
            sb.AppendLine("样本匹配测试:");

            var secondaryIndex = new Dictionary<string, DataRow>(StringComparer.OrdinalIgnoreCase);
            foreach (DataRow row in secondaryData.Rows)
            {
                var key = GetCompositeKey(row, secondaryMatchFields);
                if (!string.IsNullOrEmpty(key) && !key.Contains("__NULL__"))
                    if (!secondaryIndex.ContainsKey(key))
                        secondaryIndex.Add(key, row);
            }

            var matchCount = 0;
            foreach (DataRow primaryRow in primaryData.Rows)
            {
                var key = GetCompositeKey(primaryRow, primaryMatchFields);
                if (!string.IsNullOrEmpty(key) && !key.Contains("__NULL__"))
                {
                    if (secondaryIndex.TryGetValue(key, out _))
                    {
                        matchCount++;
                        sb.AppendLine($"成功匹配: 键={key}");
                    }
                    else
                    {
                        sb.AppendLine($"未能匹配: 键={key}");
                    }
                }
                else
                {
                    sb.AppendLine($"无效键: {key} (包含空值)");
                }
            }

            sb.AppendLine();
            sb.AppendLine(
                $"样本匹配率: {matchCount}/{primaryData.Rows.Count} ({(primaryData.Rows.Count > 0 ? (matchCount * 100.0 / primaryData.Rows.Count).ToString("F2") : "0")}%)");

            return sb.ToString();
        }
        catch (Exception ex)
        {
            return $"诊断过程出错: {ex.Message}";
        }
    }

    #endregion
}