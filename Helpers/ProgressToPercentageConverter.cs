using System.Globalization;
using System.Windows.Data;

namespace ExcelMatcher.Helpers;

/// <summary>
///     进度值到百分比转换器
/// </summary>
public class ProgressToPercentageConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is int progress && parameter is int total && total > 0)
        {
            var percentage = (double)progress / total * 100;
            return $"({percentage:F1}%)";
        }

        if (value is int progressOnly)
            // 假设最大值为100
            return $"({progressOnly:F1}%)";

        return string.Empty;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}