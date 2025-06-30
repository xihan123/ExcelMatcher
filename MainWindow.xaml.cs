using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using CommunityToolkit.Mvvm.Input;
using ExcelMatcher.ViewModels;
using MaterialDesignThemes.Wpf;

namespace ExcelMatcher;

/// <summary>
///     MainWindow.xaml 的交互逻辑
/// </summary>
public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();

        // 设置窗口样式
        SetupWindow();

        // 设置拖拽支持
        SetupDragDrop();
    }

    private void SetupWindow()
    {
        // 窗口状态变化事件
        StateChanged += MainWindow_StateChanged;

        // 处理窗口最大化时的边距
        SizeChanged += (s, e) =>
        {
            if (WindowState == WindowState.Maximized)
            {
                var margin = SystemParameters.WindowResizeBorderThickness;
                margin.Top += SystemParameters.CaptionHeight;
                Margin = margin;
            }
            else
            {
                Margin = new Thickness(0);
            }
        };
    }

    private void SetupDragDrop()
    {
        AllowDrop = true;

        // 全局文件拖拽事件（作为备用）
        DragEnter += MainWindow_DragEnter;
        DragOver += MainWindow_DragOver;
        Drop += MainWindow_Drop;
    }

    private void MainWindow_StateChanged(object sender, EventArgs e)
    {
        // 更新最大化按钮图标
        if (WindowState == WindowState.Maximized)
        {
            MaximizeButton.ToolTip = "还原";
            ((PackIcon)MaximizeButton.Content).Kind = PackIconKind.WindowRestore;
        }
        else
        {
            MaximizeButton.ToolTip = "最大化";
            ((PackIcon)MaximizeButton.Content).Kind = PackIconKind.WindowMaximize;
        }
    }

    #region Window Controls

    private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        // 允许拖动窗口
        if (e.ClickCount == 1)
            try
            {
                DragMove();
            }
            catch (InvalidOperationException)
            {
                // 忽略拖拽过程中可能出现的异常
            }
        else if (e.ClickCount == 2)
            // 双击切换最大化状态
            ToggleMaximize();
    }

    private void MinimizeButton_Click(object sender, RoutedEventArgs e)
    {
        WindowState = WindowState.Minimized;
    }

    private void MaximizeButton_Click(object sender, RoutedEventArgs e)
    {
        ToggleMaximize();
    }

    private void ToggleMaximize()
    {
        WindowState = WindowState == WindowState.Maximized
            ? WindowState.Normal
            : WindowState.Maximized;
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    #endregion

    #region Password Handling

    private void PrimaryPasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
    {
        if (sender is PasswordBox passwordBox && DataContext is MainViewModel viewModel)
            viewModel.PrimaryFilePassword = passwordBox.Password;
    }

    private void SecondaryPasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
    {
        if (sender is PasswordBox passwordBox && DataContext is MainViewModel viewModel)
            viewModel.SecondaryFilePassword = passwordBox.Password;
    }

    #endregion

    #region Global Drag Drop (Fallback)

    private void MainWindow_DragEnter(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files != null && files.Any(IsExcelFile))
                e.Effects = DragDropEffects.Copy;
            else
                e.Effects = DragDropEffects.None;
        }
        else
        {
            e.Effects = DragDropEffects.None;
        }

        e.Handled = true;
    }

    private void MainWindow_DragOver(object sender, DragEventArgs e)
    {
        MainWindow_DragEnter(sender, e);
    }

    private async void MainWindow_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            var excelFiles = files?.Where(IsExcelFile).ToArray();

            if (excelFiles != null && excelFiles.Length > 0) await HandleGlobalDroppedFiles(excelFiles);
        }

        e.Handled = true;
    }

    private async Task HandleGlobalDroppedFiles(string[] files)
    {
        if (DataContext is MainViewModel viewModel)
            try
            {
                // 智能分配文件：如果主表为空，第一个文件作为主表
                if (string.IsNullOrEmpty(viewModel.PrimaryFilePath) && files.Length > 0)
                {
                    viewModel.PrimaryFilePath = files[0];

                    if (viewModel.LoadPrimaryFileCommand is IAsyncRelayCommand asyncCommand)
                        await asyncCommand.ExecuteAsync(null);
                }

                // 如果辅助表为空且有更多文件，设置为辅助表
                if (string.IsNullOrEmpty(viewModel.SecondaryFilePath) && files.Length > 1)
                {
                    viewModel.SecondaryFilePath = files[1];

                    if (viewModel.LoadSecondaryFileCommand is IAsyncRelayCommand asyncCommand)
                        await asyncCommand.ExecuteAsync(null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"处理拖拽文件时出错: {ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
    }

    #endregion

    #region Specific Drop Zone Handlers

    // Primary File Drop Zone Events
    private void PrimaryFileDropZone_DragEnter(object sender, DragEventArgs e)
    {
        HandleFileDragEnter(e, "primary");
    }

    private void PrimaryFileDropZone_DragOver(object sender, DragEventArgs e)
    {
        HandleFileDragEnter(e, "primary");
    }

    private void PrimaryFileDropZone_DragLeave(object sender, DragEventArgs e)
    {
        // Visual feedback for drag leave can be added here
        e.Handled = true;
    }

    private async void PrimaryFileDropZone_Drop(object sender, DragEventArgs e)
    {
        await HandleFileDropAsync(e, "primary");
    }

    // Secondary File Drop Zone Events
    private void SecondaryFileDropZone_DragEnter(object sender, DragEventArgs e)
    {
        HandleFileDragEnter(e, "secondary");
    }

    private void SecondaryFileDropZone_DragOver(object sender, DragEventArgs e)
    {
        HandleFileDragEnter(e, "secondary");
    }

    private void SecondaryFileDropZone_DragLeave(object sender, DragEventArgs e)
    {
        // Visual feedback for drag leave can be added here
        e.Handled = true;
    }

    private async void SecondaryFileDropZone_Drop(object sender, DragEventArgs e)
    {
        await HandleFileDropAsync(e, "secondary");
    }

    private void HandleFileDragEnter(DragEventArgs e, string targetType)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files != null && files.Any(IsExcelFile))
                e.Effects = DragDropEffects.Copy;
            else
                e.Effects = DragDropEffects.None;
        }
        else
        {
            e.Effects = DragDropEffects.None;
        }

        e.Handled = true;
    }

    private async Task HandleFileDropAsync(DragEventArgs e, string targetType)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            var excelFile = files?.FirstOrDefault(IsExcelFile);

            if (!string.IsNullOrEmpty(excelFile) && DataContext is MainViewModel viewModel)
                try
                {
                    if (targetType == "primary")
                    {
                        viewModel.PrimaryFilePath = excelFile;

                        if (viewModel.LoadPrimaryFileCommand is IAsyncRelayCommand asyncCommand)
                            await asyncCommand.ExecuteAsync(null);
                    }
                    else if (targetType == "secondary")
                    {
                        viewModel.SecondaryFilePath = excelFile;

                        if (viewModel.LoadSecondaryFileCommand is IAsyncRelayCommand asyncCommand)
                            await asyncCommand.ExecuteAsync(null);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"处理拖拽文件时出错: {ex.Message}", "错误",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
        }

        e.Handled = true;
    }

    private static bool IsExcelFile(string filePath)
    {
        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        return extension == ".xlsx" || extension == ".xls";
    }

    private void EnablePreview_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is MainViewModel viewModel) viewModel.IsPreviewEnabled = true;
    }

    #endregion
}