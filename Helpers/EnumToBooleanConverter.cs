using System.Globalization;
using System.Windows.Data;

namespace ExcelMatcher.Helpers;

/// <summary>
///     将枚举值转换为布尔值，用于RadioButton绑定
/// </summary>
public class EnumToBooleanConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (parameter == null || value == null)
            return false;

        var parameterString = parameter.ToString();
        if (Enum.IsDefined(value.GetType(), value))
        {
            var valueString = value.ToString();
            return valueString.Equals(parameterString, StringComparison.OrdinalIgnoreCase);
        }

        return false;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (parameter == null || !(bool)value)
            return Binding.DoNothing;

        return Enum.Parse(targetType, parameter.ToString());
    }
}