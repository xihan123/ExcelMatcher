using System.Globalization;
using System.Windows.Data;

namespace ExcelMatcher.Helpers;

/// <summary>
///     反向布尔值转换器，用于反转布尔值
/// </summary>
public class InverseBooleanConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is bool boolValue) return !boolValue;
        return value;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is bool boolValue) return !boolValue;
        return value;
    }
}