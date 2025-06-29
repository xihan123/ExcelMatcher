using System.Diagnostics;
using System.IO;
using ExcelMatcher.Models;
using Newtonsoft.Json;

namespace ExcelMatcher.Services;

/// <summary>
///     配置管理服务，负责保存、加载和管理配置
/// </summary>
public class ConfigurationManager
{
    /// <summary>
    ///     保存配置到文件
    /// </summary>
    public async Task<string> SaveConfigurationAsync(Configuration configuration)
    {
        if (configuration == null)
            throw new ArgumentNullException(nameof(configuration));

        var directoryPath = GetConfigurationDirectory();
        Directory.CreateDirectory(directoryPath); // 确保目录存在

        var fileName = $"{SanitizeFileName(configuration.Name)}.json";
        var filePath = Path.Combine(directoryPath, fileName);

        // 设置序列化选项，包含详细类型信息以便正确反序列化
        var serializerSettings = new JsonSerializerSettings
        {
            Formatting = Formatting.Indented,
            NullValueHandling = NullValueHandling.Ignore,
            TypeNameHandling = TypeNameHandling.Auto,
            PreserveReferencesHandling = PreserveReferencesHandling.Objects,
            ReferenceLoopHandling = ReferenceLoopHandling.Ignore
        };

        var json = JsonConvert.SerializeObject(configuration, serializerSettings);
        await File.WriteAllTextAsync(filePath, json);

        return filePath;
    }

    /// <summary>
    ///     从文件加载配置
    /// </summary>
    public async Task<Configuration> LoadConfigurationAsync(string filePath)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"配置文件不存在: {filePath}");

        var json = await File.ReadAllTextAsync(filePath);

        // 设置反序列化选项
        var serializerSettings = new JsonSerializerSettings
        {
            TypeNameHandling = TypeNameHandling.Auto,
            PreserveReferencesHandling = PreserveReferencesHandling.Objects,
            Error = (sender, args) =>
            {
                // 处理反序列化错误
                args.ErrorContext.Handled = true;
            }
        };

        var configuration = JsonConvert.DeserializeObject<Configuration>(json, serializerSettings);
        if (configuration == null)
            throw new InvalidOperationException("无法解析配置文件");

        // 设置配置文件路径
        configuration.Path = filePath;

        // 兼容旧版配置 - 没有多工作表属性时从单工作表属性复制
        if (configuration.PrimaryWorksheets == null)
            configuration.PrimaryWorksheets = new List<string>();

        if (configuration.SecondaryWorksheets == null)
            configuration.SecondaryWorksheets = new List<string>();

        // 如果多工作表列表为空但单工作表有值，填充多工作表列表
        if (configuration.PrimaryWorksheets.Count == 0 && !string.IsNullOrEmpty(configuration.PrimaryWorksheet))
            configuration.PrimaryWorksheets.Add(configuration.PrimaryWorksheet);

        if (configuration.SecondaryWorksheets.Count == 0 && !string.IsNullOrEmpty(configuration.SecondaryWorksheet))
            configuration.SecondaryWorksheets.Add(configuration.SecondaryWorksheet);

        return configuration;
    }

    /// <summary>
    ///     获取所有保存的配置
    /// </summary>
    public async Task<List<Configuration>> GetAllConfigurationsAsync()
    {
        var directoryPath = GetConfigurationDirectory();
        if (!Directory.Exists(directoryPath))
            return new List<Configuration>();

        var configFiles = Directory.GetFiles(directoryPath, "*.json");
        var configurations = new List<Configuration>();

        foreach (var filePath in configFiles)
            try
            {
                var config = await LoadConfigurationAsync(filePath);
                configurations.Add(config);
            }
            catch (Exception)
            {
                // 跳过无法加载的配置文件
            }

        return configurations;
    }

    /// <summary>
    ///     删除配置文件
    /// </summary>
    public async Task<bool> DeleteConfigurationAsync(string configFilePath)
    {
        if (string.IsNullOrEmpty(configFilePath))
            throw new ArgumentNullException(nameof(configFilePath));

        try
        {
            if (File.Exists(configFilePath))
            {
                // 异步删除文件
                await Task.Run(() => File.Delete(configFilePath));
                return true;
            }

            return false;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"删除配置文件失败: {ex.Message}", ex);
        }
    }

    /// <summary>
    ///     打开配置文件目录
    /// </summary>
    public void OpenConfigurationDirectory()
    {
        try
        {
            var directoryPath = GetConfigurationDirectory();

            // 确保目录存在
            if (!Directory.Exists(directoryPath))
                Directory.CreateDirectory(directoryPath);

            // 使用系统默认文件管理器打开目录
            var startInfo = new ProcessStartInfo
            {
                FileName = directoryPath,
                UseShellExecute = true,
                Verb = "open"
            };

            Process.Start(startInfo);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"打开配置目录失败: {ex.Message}", ex);
        }
    }

    /// <summary>
    ///     获取配置文件存储目录
    /// </summary>
    public string GetConfigurationDirectory()
    {
        var appDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var configDirectory = Path.Combine(appDataFolder, "ExcelMatcher", "Configurations");

        // 确保目录存在
        if (!Directory.Exists(configDirectory))
            Directory.CreateDirectory(configDirectory);

        return configDirectory;
    }

    /// <summary>
    ///     净化文件名，移除不合法字符
    /// </summary>
    private string SanitizeFileName(string fileName)
    {
        var invalidChars = Path.GetInvalidFileNameChars();
        var validFileName = new string(fileName.Select(c => invalidChars.Contains(c) ? '_' : c).ToArray());
        return validFileName;
    }
}