using System.Collections;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ExcelMatcher.Views;

/// <summary>
///     FieldMappingControl.xaml 的交互逻辑
/// </summary>
public partial class FieldMappingControl : UserControl
{
    // 源字段列表依赖属性
    public static readonly DependencyProperty SourceFieldsProperty =
        DependencyProperty.Register("SourceFields", typeof(IEnumerable),
            typeof(FieldMappingControl), new PropertyMetadata(null));

    // 目标字段列表依赖属性
    public static readonly DependencyProperty TargetFieldsProperty =
        DependencyProperty.Register("TargetFields", typeof(IEnumerable),
            typeof(FieldMappingControl), new PropertyMetadata(null));

    // 字段映射列表依赖属性
    public static readonly DependencyProperty MappingsProperty =
        DependencyProperty.Register("Mappings", typeof(IEnumerable),
            typeof(FieldMappingControl), new PropertyMetadata(null));

    // 添加映射命令依赖属性
    public static readonly DependencyProperty AddMappingCommandProperty =
        DependencyProperty.Register("AddMappingCommand", typeof(ICommand),
            typeof(FieldMappingControl), new PropertyMetadata(null));

    // 移除映射命令依赖属性
    public static readonly DependencyProperty RemoveMappingCommandProperty =
        DependencyProperty.Register("RemoveMappingCommand", typeof(ICommand),
            typeof(FieldMappingControl), new PropertyMetadata(null));

    public FieldMappingControl()
    {
        InitializeComponent();
    }

    public IEnumerable SourceFields
    {
        get => (IEnumerable)GetValue(SourceFieldsProperty);
        set => SetValue(SourceFieldsProperty, value);
    }

    public IEnumerable TargetFields
    {
        get => (IEnumerable)GetValue(TargetFieldsProperty);
        set => SetValue(TargetFieldsProperty, value);
    }

    public IEnumerable Mappings
    {
        get => (IEnumerable)GetValue(MappingsProperty);
        set => SetValue(MappingsProperty, value);
    }

    public ICommand AddMappingCommand
    {
        get => (ICommand)GetValue(AddMappingCommandProperty);
        set => SetValue(AddMappingCommandProperty, value);
    }

    public ICommand RemoveMappingCommand
    {
        get => (ICommand)GetValue(RemoveMappingCommandProperty);
        set => SetValue(RemoveMappingCommandProperty, value);
    }
}