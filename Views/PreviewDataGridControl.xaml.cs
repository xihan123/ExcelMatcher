using System.Data;
using System.Windows;
using System.Windows.Controls;

namespace ExcelMatcher.Views;

/// <summary>
///     PreviewDataGridControl.xaml 的交互逻辑
/// </summary>
public partial class PreviewDataGridControl : UserControl
{
    // 标题依赖属性
    public static readonly DependencyProperty TitleProperty =
        DependencyProperty.Register("Title", typeof(string), typeof(PreviewDataGridControl),
            new PropertyMetadata("数据预览"));

    // 数据源依赖属性
    public static readonly DependencyProperty DataSourceProperty =
        DependencyProperty.Register("DataSource", typeof(DataTable), typeof(PreviewDataGridControl),
            new PropertyMetadata(null));

    // 最大高度依赖属性
    public new static readonly DependencyProperty MaxHeightProperty =
        DependencyProperty.Register("MaxHeight", typeof(double), typeof(PreviewDataGridControl),
            new PropertyMetadata(300.0));

    public PreviewDataGridControl()
    {
        InitializeComponent();
    }

    public string Title
    {
        get => (string)GetValue(TitleProperty);
        set => SetValue(TitleProperty, value);
    }

    public DataTable DataSource
    {
        get => (DataTable)GetValue(DataSourceProperty);
        set => SetValue(DataSourceProperty, value);
    }

    public new double MaxHeight
    {
        get => (double)GetValue(MaxHeightProperty);
        set => SetValue(MaxHeightProperty, value);
    }
}