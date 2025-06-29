using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace ExcelMatcher.Helpers;

/// <summary>
///     多项显示可见性转换器
/// </summary>
public class MultipleItemsToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is int count) return count > 1 ? Visibility.Visible : Visibility.Collapsed;
        return Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}