using NLog;

namespace ExcelMatcher.Helpers;

/// <summary>
///     日志工具类，提供全局日志记录功能
/// </summary>
public static class Logger
{
    private static readonly NLog.Logger _logger = LogManager.GetCurrentClassLogger();

    /// <summary>
    ///     记录调试信息
    /// </summary>
    /// <param name="message">日志信息</param>
    public static void Debug(string message)
    {
        _logger.Debug(message);
    }

    /// <summary>
    ///     记录一般信息
    /// </summary>
    /// <param name="message">日志信息</param>
    public static void Info(string message)
    {
        _logger.Info(message);
    }

    /// <summary>
    ///     记录警告信息
    /// </summary>
    /// <param name="message">日志信息</param>
    public static void Warning(string message)
    {
        _logger.Warn(message);
    }

    /// <summary>
    ///     记录错误信息
    /// </summary>
    /// <param name="message">日志信息</param>
    /// <param name="exception">异常对象</param>
    public static void Error(string message, Exception? exception = null)
    {
        if (exception == null)
            _logger.Error(message);
        else
            _logger.Error(exception, message);
    }

    /// <summary>
    ///     记录致命错误信息
    /// </summary>
    /// <param name="message">日志信息</param>
    /// <param name="exception">异常对象</param>
    public static void Fatal(string message, Exception? exception = null)
    {
        if (exception == null)
            _logger.Fatal(message);
        else
            _logger.Fatal(exception, message);
    }
}