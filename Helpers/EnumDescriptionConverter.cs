using System.Globalization;
using System.Windows.Data;

namespace ExcelMatcher.Helpers;

/// <summary>
///     枚举描述转换器，将枚举值转换为其描述文本
/// </summary>
public class EnumDescriptionConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is Enum enumValue) return EnumHelper.GetDescription(enumValue);
        return value.ToString();
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is string stringValue && targetType.IsEnum)
        {
            var result = EnumHelper.GetValueFromDescription<Enum>(stringValue);
            if (result != null) return result;
            return Enum.Parse(targetType, stringValue);
        }

        return value;
    }
}