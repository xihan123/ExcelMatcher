using System.Globalization;
using System.Windows.Data;

namespace ExcelMatcher.Helpers;

/// <summary>
///     布尔值到透明度转换器
/// </summary>
public class BoolToOpacityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is bool boolValue) return boolValue ? 1.0 : 0.5;
        return 1.0;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}