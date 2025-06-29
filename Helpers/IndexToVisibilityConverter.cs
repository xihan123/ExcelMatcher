using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace ExcelMatcher.Helpers;

/// <summary>
///     索引到可见性转换器，用于在ItemsControl中根据索引决定元素可见性
/// </summary>
public class IndexToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is ContentPresenter contentPresenter)
        {
            var itemsControl = ItemsControl.ItemsControlFromItemContainer(contentPresenter);
            if (itemsControl != null)
            {
                var index = itemsControl.ItemContainerGenerator.IndexFromContainer(contentPresenter);

                // 第一个元素不显示逻辑运算符
                return index == 0 ? Visibility.Collapsed : Visibility.Visible;
            }
        }

        return Visibility.Visible;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}