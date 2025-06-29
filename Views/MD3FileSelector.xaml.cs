using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;

namespace ExcelMatcher.Views;

/// <summary>
///     MD3FileSelector.xaml 的交互逻辑
/// </summary>
public partial class MD3FileSelector : UserControl
{
    public static readonly DependencyProperty FilePathProperty =
        DependencyProperty.Register(nameof(FilePath), typeof(string), typeof(MD3FileSelector),
            new FrameworkPropertyMetadata(string.Empty, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault,
                OnFilePathChanged));

    public static readonly DependencyProperty PasswordProperty =
        DependencyProperty.Register(nameof(Password), typeof(string), typeof(MD3FileSelector),
            new FrameworkPropertyMetadata(string.Empty, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));

    public static readonly DependencyProperty PlaceholderProperty =
        DependencyProperty.Register(nameof(Placeholder), typeof(string), typeof(MD3FileSelector),
            new PropertyMetadata("选择Excel文件"));

    public static readonly DependencyProperty HasFileProperty =
        DependencyProperty.Register(nameof(HasFile), typeof(bool), typeof(MD3FileSelector),
            new PropertyMetadata(false));

    public static readonly DependencyProperty IsDragOverProperty =
        DependencyProperty.Register(nameof(IsDragOver), typeof(bool), typeof(MD3FileSelector),
            new PropertyMetadata(false));

    public static readonly RoutedEvent FileSelectedEvent =
        EventManager.RegisterRoutedEvent(nameof(FileSelected), RoutingStrategy.Bubble,
            typeof(RoutedEventHandler), typeof(MD3FileSelector));

    public MD3FileSelector()
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

    public string Placeholder
    {
        get => (string)GetValue(PlaceholderProperty);
        set => SetValue(PlaceholderProperty, value);
    }

    public bool HasFile
    {
        get => (bool)GetValue(HasFileProperty);
        private set => SetValue(HasFileProperty, value);
    }

    public bool IsDragOver
    {
        get => (bool)GetValue(IsDragOverProperty);
        private set => SetValue(IsDragOverProperty, value);
    }

    public event RoutedEventHandler FileSelected
    {
        add => AddHandler(FileSelectedEvent, value);
        remove => RemoveHandler(FileSelectedEvent, value);
    }

    private static void OnFilePathChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is MD3FileSelector selector) selector.HasFile = !string.IsNullOrEmpty((string)e.NewValue);
    }

    private void BrowseButton_Click(object sender, RoutedEventArgs e)
    {
        var openFileDialog = new OpenFileDialog
        {
            Filter = "Excel文件|*.xlsx;*.xls|所有文件|*.*",
            Title = "选择Excel文件"
        };

        if (openFileDialog.ShowDialog() == true)
        {
            FilePath = openFileDialog.FileName;
            RaiseEvent(new RoutedEventArgs(FileSelectedEvent));
        }
    }

    private void Border_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (!HasFile) BrowseButton_Click(sender, e);
    }

    private void Border_DragEnter(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files != null && files.Any(IsExcelFile))
            {
                e.Effects = DragDropEffects.Copy;
                IsDragOver = true;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }
        else
        {
            e.Effects = DragDropEffects.None;
        }

        e.Handled = true;
    }

    private void Border_DragOver(object sender, DragEventArgs e)
    {
        Border_DragEnter(sender, e);
    }

    private void Border_DragLeave(object sender, DragEventArgs e)
    {
        IsDragOver = false;
        e.Handled = true;
    }

    private void Border_Drop(object sender, DragEventArgs e)
    {
        IsDragOver = false;

        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            var excelFile = files?.FirstOrDefault(IsExcelFile);

            if (!string.IsNullOrEmpty(excelFile))
            {
                FilePath = excelFile;
                RaiseEvent(new RoutedEventArgs(FileSelectedEvent));
            }
        }

        e.Handled = true;
    }

    private void PasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
    {
        if (sender is PasswordBox passwordBox) Password = passwordBox.Password;
    }

    private static bool IsExcelFile(string filePath)
    {
        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        return extension == ".xlsx" || extension == ".xls";
    }
}