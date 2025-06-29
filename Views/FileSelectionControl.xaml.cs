using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ExcelMatcher.Views;

/// <summary>
///     FileSelectionControl.xaml 的交互逻辑
/// </summary>
public partial class FileSelectionControl : UserControl
{
    // 文件路径依赖属性
    public static readonly DependencyProperty FilePathProperty =
        DependencyProperty.Register("FilePath", typeof(string), typeof(FileSelectionControl),
            new PropertyMetadata(string.Empty));

    // 文件密码依赖属性
    public static readonly DependencyProperty PasswordProperty =
        DependencyProperty.Register("Password", typeof(string), typeof(FileSelectionControl),
            new PropertyMetadata(string.Empty));

    // 标题依赖属性
    public static readonly DependencyProperty TitleProperty =
        DependencyProperty.Register("Title", typeof(string), typeof(FileSelectionControl),
            new PropertyMetadata("选择文件"));

    // 浏览命令依赖属性
    public static readonly DependencyProperty BrowseCommandProperty =
        DependencyProperty.Register("BrowseCommand", typeof(ICommand),
            typeof(FileSelectionControl), new PropertyMetadata(null));

    public FileSelectionControl()
    {
        InitializeComponent();
    }

    public string FilePath
    {
        get => (string)GetValue(FilePathProperty);
        set => SetValue(FilePathProperty, value);
    }

    public string Password
    {
        get => (string)GetValue(PasswordProperty);
        set => SetValue(PasswordProperty, value);
    }

    public string Title
    {
        get => (string)GetValue(TitleProperty);
        set => SetValue(TitleProperty, value);
    }

    public ICommand BrowseCommand
    {
        get => (ICommand)GetValue(BrowseCommandProperty);
        set => SetValue(BrowseCommandProperty, value);
    }
}