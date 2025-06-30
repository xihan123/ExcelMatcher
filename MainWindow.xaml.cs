using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using CommunityToolkit.Mvvm.Input;
using ExcelMatcher.ViewModels;
using MaterialDesignThemes.Wpf;

namespace ExcelMatcher;

/// <summary>
///     MainWindow.xaml 的交互逻辑
/// </summary>
public partial class MainWindow : Window
{
    private bool _isDragOver;

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

        // 全局文件拖拽事件
        DragEnter += MainWindow_DragEnter;
        DragOver += MainWindow_DragOver;
        DragLeave += MainWindow_DragLeave;
        Drop += MainWindow_Drop;
    }

    private void MainWindow_StateChanged(object? sender, EventArgs e)
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

    #region Global Drag Drop (Enhanced)

    private void MainWindow_DragEnter(object sender, DragEventArgs e)
    {
        if (DataContext is MainViewModel viewModel) viewModel.IsDragOver = true;
        _isDragOver = true;
        HandleGlobalDragEnter(e);
    }

    private void MainWindow_DragOver(object sender, DragEventArgs e)
    {
        HandleGlobalDragEnter(e);
    }

    private void MainWindow_DragLeave(object sender, DragEventArgs e)
    {
        if (DataContext is MainViewModel viewModel) viewModel.IsDragOver = false;
        _isDragOver = false;
        // 移除拖拽视觉反馈
        RemoveDragOverEffect();
        e.Handled = true;
    }

    private void HandleGlobalDragEnter(DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            var excelFiles = files?.Where(IsExcelFile).ToArray();

            if (excelFiles != null && excelFiles.Length > 0)
            {
                e.Effects = DragDropEffects.Copy;
                // 添加拖拽视觉反馈
                AddDragOverEffect();

                // 显示拖拽提示
                ShowDragHint(excelFiles);
            }
            else
            {
                e.Effects = DragDropEffects.None;
                RemoveDragOverEffect();
            }
        }
        else
        {
            e.Effects = DragDropEffects.None;
            RemoveDragOverEffect();
        }

        e.Handled = true;
    }

    private async void MainWindow_Drop(object sender, DragEventArgs e)
    {
        if (DataContext is MainViewModel viewModel) viewModel.IsDragOver = false;
        _isDragOver = false;
        RemoveDragOverEffect();

        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            var excelFiles = files?.Where(IsExcelFile).ToArray();

            if (excelFiles != null && excelFiles.Length > 0)
                await HandleGlobalDroppedFiles(excelFiles);
        }

        e.Handled = true;
    }

    private async Task HandleGlobalDroppedFiles(string[] files)
    {
        if (DataContext is not MainViewModel viewModel)
            return;

        try
        {
            var primaryEmpty = string.IsNullOrEmpty(viewModel.PrimaryFilePath);
            var secondaryEmpty = string.IsNullOrEmpty(viewModel.SecondaryFilePath);

            if (files.Length == 1)
                // 单个文件的处理逻辑
                await HandleSingleFileAssignment(files[0], viewModel, primaryEmpty, secondaryEmpty);
            else if (files.Length > 1)
                // 多个文件的处理逻辑
                await HandleMultipleFileAssignment(files, viewModel, primaryEmpty, secondaryEmpty);
        }
        catch (Exception ex)
        {
            await ShowErrorMessage($"处理拖拽文件时出错: {ex.Message}");
        }
    }

    private async Task HandleSingleFileAssignment(string filePath, MainViewModel viewModel, bool primaryEmpty,
        bool secondaryEmpty)
    {
        if (primaryEmpty)
        {
            // 主表为空，直接设置为主表
            await SetPrimaryFile(filePath, viewModel);
            await ShowSuccessMessage($"已将文件设置为主表: {Path.GetFileName(filePath)}");
        }
        else if (secondaryEmpty)
        {
            // 主表有文件但辅助表为空，设置为辅助表
            await SetSecondaryFile(filePath, viewModel);
            await ShowSuccessMessage($"已将文件设置为辅助表: {Path.GetFileName(filePath)}");
        }
        else
        {
            // 两个表都有文件，让用户选择替换哪个
            await ShowFileReplaceDialog(filePath, viewModel);
        }
    }

    private async Task HandleMultipleFileAssignment(string[] files, MainViewModel viewModel, bool primaryEmpty,
        bool secondaryEmpty)
    {
        if (primaryEmpty && secondaryEmpty)
        {
            // 两个表都为空，第一个设为主表，第二个设为辅助表
            await SetPrimaryFile(files[0], viewModel);
            if (files.Length > 1)
                await SetSecondaryFile(files[1], viewModel);

            await ShowSuccessMessage($"已设置主表: {Path.GetFileName(files[0])}" +
                                     (files.Length > 1 ? $"\n辅助表: {Path.GetFileName(files[1])}" : ""));
        }
        else if (primaryEmpty)
        {
            // 只有主表为空，设置第一个文件为主表
            await SetPrimaryFile(files[0], viewModel);
            await ShowSuccessMessage($"已将文件设置为主表: {Path.GetFileName(files[0])}");

            if (files.Length > 1)
                await ShowMultipleFilesDialog(files.Skip(1).ToArray(), viewModel, false, secondaryEmpty);
        }
        else if (secondaryEmpty)
        {
            // 只有辅助表为空，设置第一个文件为辅助表
            await SetSecondaryFile(files[0], viewModel);
            await ShowSuccessMessage($"已将文件设置为辅助表: {Path.GetFileName(files[0])}");

            if (files.Length > 1)
                await ShowMultipleFilesDialog(files.Skip(1).ToArray(), viewModel, true, false);
        }
        else
        {
            // 两个表都有文件，显示多文件处理对话框
            await ShowMultipleFilesDialog(files, viewModel, true, true);
        }
    }

    private async Task SetPrimaryFile(string filePath, MainViewModel viewModel)
    {
        viewModel.PrimaryFilePath = filePath;
        if (viewModel.LoadPrimaryFileCommand is IAsyncRelayCommand asyncCommand)
            await asyncCommand.ExecuteAsync(null);
    }

    private async Task SetSecondaryFile(string filePath, MainViewModel viewModel)
    {
        viewModel.SecondaryFilePath = filePath;
        if (viewModel.LoadSecondaryFileCommand is IAsyncRelayCommand asyncCommand)
            await asyncCommand.ExecuteAsync(null);
    }

    #endregion

    #region Dialog Methods

    private async Task ShowFileReplaceDialog(string filePath, MainViewModel viewModel)
    {
        var dialogContent = new StackPanel { Margin = new Thickness(24) };

        // 标题
        dialogContent.Children.Add(new TextBlock
        {
            Text = "选择文件位置",
            Style = FindResource("MD3TitleLarge") as Style,
            Margin = new Thickness(0, 0, 0, 16)
        });

        // 说明
        dialogContent.Children.Add(new TextBlock
        {
            Text = $"主表和辅助表都已有文件，请选择将 '{Path.GetFileName(filePath)}' 设置为:",
            Style = FindResource("MD3BodyLarge") as Style,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 24)
        });

        // 当前文件信息
        var currentFilesCard = new Card
        {
            Style = FindResource("MD3OutlinedCard") as Style,
            Margin = new Thickness(0, 0, 0, 24)
        };

        var currentFilesPanel = new StackPanel();
        currentFilesPanel.Children.Add(new TextBlock
        {
            Text = "当前文件:",
            Style = FindResource("MD3LabelLarge") as Style,
            Margin = new Thickness(0, 0, 0, 8)
        });

        currentFilesPanel.Children.Add(new TextBlock
        {
            Text = $"主表: {Path.GetFileName(viewModel.PrimaryFilePath)}",
            Style = FindResource("MD3BodyMedium") as Style,
            Margin = new Thickness(0, 0, 0, 4)
        });

        currentFilesPanel.Children.Add(new TextBlock
        {
            Text = $"辅助表: {Path.GetFileName(viewModel.SecondaryFilePath)}",
            Style = FindResource("MD3BodyMedium") as Style
        });

        currentFilesCard.Content = currentFilesPanel;
        dialogContent.Children.Add(currentFilesCard);

        // 按钮
        var buttonPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right
        };

        var cancelButton = new Button
        {
            Content = "取消",
            Style = FindResource("MD3OutlinedButton") as Style,
            Margin = new Thickness(0, 0, 8, 0)
        };

        var primaryButton = new Button
        {
            Content = "替换主表",
            Style = FindResource("MD3FilledButton") as Style,
            Margin = new Thickness(0, 0, 8, 0)
        };

        var secondaryButton = new Button
        {
            Content = "替换辅助表",
            Style = FindResource("MD3FilledButton") as Style
        };

        buttonPanel.Children.Add(cancelButton);
        buttonPanel.Children.Add(primaryButton);
        buttonPanel.Children.Add(secondaryButton);
        dialogContent.Children.Add(buttonPanel);

        // 事件处理
        var dialogResult = "";
        cancelButton.Click += (s, e) => { DialogHost.Close("RootDialog"); };
        primaryButton.Click += (s, e) =>
        {
            dialogResult = "primary";
            DialogHost.Close("RootDialog");
        };
        secondaryButton.Click += (s, e) =>
        {
            dialogResult = "secondary";
            DialogHost.Close("RootDialog");
        };

        await DialogHost.Show(dialogContent, "RootDialog");

        // 处理结果
        if (dialogResult == "primary")
        {
            await SetPrimaryFile(filePath, viewModel);
            await ShowSuccessMessage($"已替换主表文件: {Path.GetFileName(filePath)}");
        }
        else if (dialogResult == "secondary")
        {
            await SetSecondaryFile(filePath, viewModel);
            await ShowSuccessMessage($"已替换辅助表文件: {Path.GetFileName(filePath)}");
        }
    }

    private async Task ShowMultipleFilesDialog(string[] files, MainViewModel viewModel, bool primaryOccupied,
        bool secondaryOccupied)
    {
        var dialogContent = new StackPanel { Margin = new Thickness(24) };

        dialogContent.Children.Add(new TextBlock
        {
            Text = "处理多个文件",
            Style = FindResource("MD3TitleLarge") as Style,
            Margin = new Thickness(0, 0, 0, 16)
        });

        dialogContent.Children.Add(new TextBlock
        {
            Text = $"检测到 {files.Length} 个Excel文件，请选择处理方式:",
            Style = FindResource("MD3BodyLarge") as Style,
            Margin = new Thickness(0, 0, 0, 16)
        });

        // 文件列表
        var fileListCard = new Card
        {
            Style = FindResource("MD3OutlinedCard") as Style,
            Margin = new Thickness(0, 0, 0, 24)
        };

        var fileListPanel = new StackPanel();
        foreach (var file in files.Take(5)) // 最多显示5个文件
            fileListPanel.Children.Add(new TextBlock
            {
                Text = $"• {Path.GetFileName(file)}",
                Style = FindResource("MD3BodyMedium") as Style,
                Margin = new Thickness(8, 2, 0, 2)
            });

        if (files.Length > 5)
            fileListPanel.Children.Add(new TextBlock
            {
                Text = $"... 还有 {files.Length - 5} 个文件",
                Style = FindResource("MD3BodySmall") as Style,
                Margin = new Thickness(8, 2, 0, 2),
                FontStyle = FontStyles.Italic
            });

        fileListCard.Content = fileListPanel;
        dialogContent.Children.Add(fileListCard);

        // 选项按钮
        var optionsPanel = new StackPanel { Margin = new Thickness(0, 0, 0, 24) };

        if (!primaryOccupied)
        {
            var primaryOption = new Button
            {
                Content = $"设置第一个文件为主表 ({Path.GetFileName(files[0])})",
                Style = FindResource("MD3OutlinedButton") as Style,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                Margin = new Thickness(0, 0, 0, 8)
            };
            primaryOption.Click += async (s, e) =>
            {
                DialogHost.Close("RootDialog");
                await SetPrimaryFile(files[0], viewModel);
                await ShowSuccessMessage($"已设置主表: {Path.GetFileName(files[0])}");
            };
            optionsPanel.Children.Add(primaryOption);
        }

        if (!secondaryOccupied)
        {
            var secondaryOption = new Button
            {
                Content = $"设置第一个文件为辅助表 ({Path.GetFileName(files[0])})",
                Style = FindResource("MD3OutlinedButton") as Style,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                Margin = new Thickness(0, 0, 0, 8)
            };
            secondaryOption.Click += async (s, e) =>
            {
                DialogHost.Close("RootDialog");
                await SetSecondaryFile(files[0], viewModel);
                await ShowSuccessMessage($"已设置辅助表: {Path.GetFileName(files[0])}");
            };
            optionsPanel.Children.Add(secondaryOption);
        }

        dialogContent.Children.Add(optionsPanel);

        // 取消按钮
        var cancelButton = new Button
        {
            Content = "取消",
            Style = FindResource("MD3TextButton") as Style,
            HorizontalAlignment = HorizontalAlignment.Right
        };
        cancelButton.Click += (s, e) => { DialogHost.Close("RootDialog"); };
        dialogContent.Children.Add(cancelButton);

        await DialogHost.Show(dialogContent, "RootDialog");
    }

    private async Task ShowSuccessMessage(string message)
    {
        var snackbar = MainSnackbar;
        if (snackbar?.MessageQueue != null)
            snackbar.MessageQueue.Enqueue(message, "确定", _ => { }, null, false, true, TimeSpan.FromSeconds(3));
        await Task.Delay(100); // 确保消息显示
    }

    private async Task ShowErrorMessage(string message)
    {
        var dialogContent = new StackPanel { Margin = new Thickness(24) };

        dialogContent.Children.Add(new TextBlock
        {
            Text = "错误",
            Style = FindResource("MD3TitleLarge") as Style,
            Foreground = FindResource("MD3ErrorBrush") as Brush,
            Margin = new Thickness(0, 0, 0, 16)
        });

        dialogContent.Children.Add(new TextBlock
        {
            Text = message,
            Style = FindResource("MD3BodyLarge") as Style,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 24)
        });

        var okButton = new Button
        {
            Content = "确定",
            Style = FindResource("MD3FilledButton") as Style,
            HorizontalAlignment = HorizontalAlignment.Right,
            IsDefault = true
        };
        okButton.Click += (s, e) => { DialogHost.Close("RootDialog"); };
        dialogContent.Children.Add(okButton);

        await DialogHost.Show(dialogContent, "RootDialog");
    }

    #endregion

    #region Visual Feedback

    private void AddDragOverEffect()
    {
        // 添加拖拽时的视觉效果
        if (!_isDragOver) return;

        // 可以在这里添加边框高亮或其他视觉效果
        Background = FindResource("MD3PrimaryContainerBrush") as Brush ?? Background;
        Opacity = 0.95;
    }

    private void RemoveDragOverEffect()
    {
        // 移除拖拽视觉效果
        Background = FindResource("MD3BackgroundBrush") as Brush ?? Background;
        Opacity = 1.0;
    }

    private void ShowDragHint(string[] files)
    {
        if (DataContext is not MainViewModel viewModel) return;

        var primaryEmpty = string.IsNullOrEmpty(viewModel.PrimaryFilePath);
        var secondaryEmpty = string.IsNullOrEmpty(viewModel.SecondaryFilePath);

        var hint = files.Length switch
        {
            1 when primaryEmpty => $"将设置为主表: {Path.GetFileName(files[0])}",
            1 when secondaryEmpty => $"将设置为辅助表: {Path.GetFileName(files[0])}",
            1 => $"选择替换主表或辅助表: {Path.GetFileName(files[0])}",
            > 1 when primaryEmpty && secondaryEmpty => $"将设置 {files.Length} 个文件",
            > 1 => $"处理 {files.Length} 个Excel文件",
            _ => "拖放文件到此处"
        };

        // 这里可以显示临时提示，或者更新状态栏
        // 由于需要避免阻塞UI，这里只是设置一个简单的状态
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
                        await SetPrimaryFile(excelFile, viewModel);
                        await ShowSuccessMessage($"已设置主表: {Path.GetFileName(excelFile)}");
                    }
                    else if (targetType == "secondary")
                    {
                        await SetSecondaryFile(excelFile, viewModel);
                        await ShowSuccessMessage($"已设置辅助表: {Path.GetFileName(excelFile)}");
                    }
                }
                catch (Exception ex)
                {
                    await ShowErrorMessage($"处理拖拽文件时出错: {ex.Message}");
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
        if (DataContext is MainViewModel viewModel)
            viewModel.IsPreviewEnabled = true;
    }

    #endregion
}