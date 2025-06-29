using System.Collections;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ExcelMatcher.Models;

namespace ExcelMatcher.Views;

/// <summary>
///     DataFilterControl.xaml 的交互逻辑
/// </summary>
public partial class DataFilterControl : UserControl
{
    // 标题依赖属性
    public static readonly DependencyProperty TitleProperty =
        DependencyProperty.Register("Title", typeof(string), typeof(DataFilterControl),
            new PropertyMetadata("筛选条件"));

    // 可用字段依赖属性
    public static readonly DependencyProperty AvailableFieldsProperty =
        DependencyProperty.Register("AvailableFields", typeof(IEnumerable),
            typeof(DataFilterControl), new PropertyMetadata(null));

    // 筛选条件依赖属性
    public static readonly DependencyProperty FiltersProperty =
        DependencyProperty.Register("Filters", typeof(IEnumerable),
            typeof(DataFilterControl), new PropertyMetadata(null));

    // 筛选操作符依赖属性
    public static readonly DependencyProperty FilterOperatorsProperty =
        DependencyProperty.Register("FilterOperators", typeof(IEnumerable<FilterOperator>),
            typeof(DataFilterControl), new PropertyMetadata(null));

    // 逻辑操作符依赖属性
    public static readonly DependencyProperty LogicalOperatorsProperty =
        DependencyProperty.Register("LogicalOperators", typeof(IEnumerable<LogicalOperator>),
            typeof(DataFilterControl), new PropertyMetadata(null));

    // 添加筛选命令依赖属性
    public static readonly DependencyProperty AddFilterCommandProperty =
        DependencyProperty.Register("AddFilterCommand", typeof(ICommand),
            typeof(DataFilterControl), new PropertyMetadata(null));

    // 移除筛选命令依赖属性
    public static readonly DependencyProperty RemoveFilterCommandProperty =
        DependencyProperty.Register("RemoveFilterCommand", typeof(ICommand),
            typeof(DataFilterControl), new PropertyMetadata(null));

    public DataFilterControl()
    {
        InitializeComponent();
    }

    public string Title
    {
        get => (string)GetValue(TitleProperty);
        set => SetValue(TitleProperty, value);
    }

    public IEnumerable AvailableFields
    {
        get => (IEnumerable)GetValue(AvailableFieldsProperty);
        set => SetValue(AvailableFieldsProperty, value);
    }

    public IEnumerable Filters
    {
        get => (IEnumerable)GetValue(FiltersProperty);
        set => SetValue(FiltersProperty, value);
    }

    public IEnumerable<FilterOperator> FilterOperators
    {
        get => (IEnumerable<FilterOperator>)GetValue(FilterOperatorsProperty);
        set => SetValue(FilterOperatorsProperty, value);
    }

    public IEnumerable<LogicalOperator> LogicalOperators
    {
        get => (IEnumerable<LogicalOperator>)GetValue(LogicalOperatorsProperty);
        set => SetValue(LogicalOperatorsProperty, value);
    }

    public ICommand AddFilterCommand
    {
        get => (ICommand)GetValue(AddFilterCommandProperty);
        set => SetValue(AddFilterCommandProperty, value);
    }

    public ICommand RemoveFilterCommand
    {
        get => (ICommand)GetValue(RemoveFilterCommandProperty);
        set => SetValue(RemoveFilterCommandProperty, value);
    }
}