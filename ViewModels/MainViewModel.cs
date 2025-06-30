using System.Collections;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using CommunityToolkit.Mvvm.Input;
using ExcelMatcher.Models;
using ExcelMatcher.Services;
using MaterialDesignThemes.Wpf;
using Microsoft.Win32;

namespace ExcelMatcher.ViewModels;

/// <summary>
///     主窗体视图模型
/// </summary>
public class MainViewModel : BaseViewModel
{
    // 命令字段
    private readonly RelayCommand _addFieldMappingCommand;
    private readonly RelayCommand _addPrimaryFilterCommand;
    private readonly RelayCommand _addSecondaryFilterCommand;
    private readonly AsyncRelayCommand _browsePrimaryFileCommand;
    private readonly AsyncRelayCommand _browseSecondaryFileCommand;
    private readonly ConfigurationManager _configurationManager;
    private readonly AsyncRelayCommand<Configuration> _deleteConfigurationCommand;

    /// <summary>
    ///     诊断匹配问题命令
    /// </summary>
    private readonly AsyncRelayCommand _diagnoseMatchingCommand;

    private readonly ExcelFileManager _excelFileManager;
    private readonly AsyncRelayCommand _loadConfigurationCommand;
    private readonly AsyncRelayCommand _loadPrimaryFileCommand;
    private readonly AsyncRelayCommand _loadSecondaryFileCommand;
    private readonly RelayCommand _openConfigurationDirectoryCommand;

    /// <summary>
    ///     刷新Excel文件数据
    /// </summary>
    private readonly AsyncRelayCommand _refreshDataCommand;

    private readonly RelayCommand<FieldMapping> _removeFieldMappingCommand;
    private readonly RelayCommand<FilterCondition> _removePrimaryFilterCommand;
    private readonly RelayCommand<FilterCondition> _removeSecondaryFilterCommand;
    private readonly RelayCommand _resetConfigurationCommand;
    private readonly AsyncRelayCommand _saveConfigurationCommand;
    private readonly AsyncRelayCommand _startMergeCommand;
    private readonly RelayCommand<IList> _updateSelectedPrimaryMatchFieldsCommand;
    private readonly RelayCommand<IList> _updateSelectedPrimaryWorksheetsCommand;
    private readonly RelayCommand<IList> _updateSelectedSecondaryMatchFieldsCommand;
    private readonly RelayCommand<IList> _updateSelectedSecondaryWorksheetsCommand;
    private ObservableCollection<FieldMapping> _fieldMappings;
    private bool _isBusy;
    private ObservableCollection<string> _primaryColumns;

    // 主表文件相关属性
    private ExcelFile _primaryFile = new();
    private string _primaryFilePassword = string.Empty;
    private string _primaryFilePath = string.Empty;

    // 数据筛选相关属性
    private ObservableCollection<FilterCondition> _primaryFilters;
    private DataTable? _primaryPreviewData;
    private ObservableCollection<string> _primaryWorksheets;
    private int _progressMaximum = 100;


    // 进度和状态相关属性
    private int _progressValue;
    private ObservableCollection<string> _secondaryColumns;

    // 辅助表文件相关属性
    private ExcelFile _secondaryFile = new();
    private string _secondaryFilePassword = string.Empty;
    private string _secondaryFilePath = string.Empty;
    private ObservableCollection<FilterCondition> _secondaryFilters;
    private DataTable? _secondaryPreviewData;
    private ObservableCollection<string> _secondaryWorksheets;

    // 字段匹配相关属性
    private ObservableCollection<string> _selectedPrimaryMatchFields;
    private string _selectedPrimaryWorksheet = string.Empty;
    private ObservableCollection<string> _selectedPrimaryWorksheets;
    private ObservableCollection<string> _selectedSecondaryMatchFields;
    private string _selectedSecondaryWorksheet = string.Empty;
    private ObservableCollection<string> _selectedSecondaryWorksheets;
    private string _statusMessage = "准备就绪";

    // 预览开关属性
    private bool _isPreviewEnabled = true;

    public bool IsPreviewEnabled
    {
        get => _isPreviewEnabled;
        set
        {
            if (SetProperty(ref _isPreviewEnabled, value))
            {
                // 当预览被禁用时清空预览数据
                if (!value)
                {
                    PrimaryPreviewData = null;
                    SecondaryPreviewData = null;
                }
                // 当预览被启用时重新加载预览数据
                else if (PrimaryFile?.IsLoaded == true && SecondaryFile?.IsLoaded == true)
                {
                    LoadPreviewDataWithFiltersAsync().ConfigureAwait(false);
                }
            }
        }
    }

    // 构造函数
    public MainViewModel(ExcelFileManager excelFileManager, ConfigurationManager configurationManager)
    {
        _excelFileManager = excelFileManager ?? throw new ArgumentNullException(nameof(excelFileManager));
        _configurationManager = configurationManager ?? throw new ArgumentNullException(nameof(configurationManager));

        // 初始化集合
        _primaryWorksheets = new ObservableCollection<string>();
        _secondaryWorksheets = new ObservableCollection<string>();
        _primaryColumns = new ObservableCollection<string>();
        _secondaryColumns = new ObservableCollection<string>();
        _selectedPrimaryMatchFields = new ObservableCollection<string>();
        _selectedSecondaryMatchFields = new ObservableCollection<string>();
        _selectedPrimaryWorksheets = new ObservableCollection<string>();
        _selectedSecondaryWorksheets = new ObservableCollection<string>();
        _fieldMappings = new ObservableCollection<FieldMapping>();
        _primaryFilters = new ObservableCollection<FilterCondition>();
        _secondaryFilters = new ObservableCollection<FilterCondition>();

        // 初始化命令
        _browsePrimaryFileCommand = new AsyncRelayCommand(BrowsePrimaryFileAsync);
        _browseSecondaryFileCommand = new AsyncRelayCommand(BrowseSecondaryFileAsync);
        _loadPrimaryFileCommand = new AsyncRelayCommand(() => LoadPrimaryFileAsync());
        _loadSecondaryFileCommand = new AsyncRelayCommand(() => LoadSecondaryFileAsync());
        _addFieldMappingCommand = new RelayCommand(AddFieldMapping, CanAddFieldMapping);
        _refreshDataCommand = new AsyncRelayCommand(RefreshDataAsync, CanRefreshData);
        _diagnoseMatchingCommand = new AsyncRelayCommand(DiagnoseMatchingAsync, CanDiagnoseMatching);
        _removeFieldMappingCommand = new RelayCommand<FieldMapping>(RemoveFieldMapping);
        _addPrimaryFilterCommand = new RelayCommand(AddPrimaryFilter, CanAddPrimaryFilter);
        _removePrimaryFilterCommand = new RelayCommand<FilterCondition>(RemovePrimaryFilter);
        _addSecondaryFilterCommand = new RelayCommand(AddSecondaryFilter, CanAddSecondaryFilter);
        _removeSecondaryFilterCommand = new RelayCommand<FilterCondition>(RemoveSecondaryFilter);
        _startMergeCommand = new AsyncRelayCommand(StartMergeAsync, CanStartMerge);
        _deleteConfigurationCommand = new AsyncRelayCommand<Configuration>(DeleteConfigurationAsync);
        _openConfigurationDirectoryCommand = new RelayCommand(OpenConfigurationDirectory);
        _saveConfigurationCommand = new AsyncRelayCommand(SaveConfigurationAsync);
        _loadConfigurationCommand = new AsyncRelayCommand(LoadConfigurationAsync);
        _resetConfigurationCommand = new RelayCommand(ResetConfiguration);

        _updateSelectedPrimaryMatchFieldsCommand = new RelayCommand<IList>(UpdateSelectedPrimaryMatchFields);
        _updateSelectedSecondaryMatchFieldsCommand = new RelayCommand<IList>(UpdateSelectedSecondaryMatchFields);
        _updateSelectedPrimaryWorksheetsCommand = new RelayCommand<IList>(UpdateSelectedPrimaryWorksheets);
        _updateSelectedSecondaryWorksheetsCommand = new RelayCommand<IList>(UpdateSelectedSecondaryWorksheets);

        // 添加筛选条件变更监听，用于自动刷新预览
        _primaryFilters.CollectionChanged += FilterConditionsCollectionChanged;
        _secondaryFilters.CollectionChanged += FilterConditionsCollectionChanged;
    }

    public ICommand RefreshDataCommand => _refreshDataCommand;
    public ICommand DiagnoseMatchingCommand => _diagnoseMatchingCommand;

    #region 属性

    // 主表文件属性
    public ExcelFile PrimaryFile
    {
        get => _primaryFile;
        set => SetProperty(ref _primaryFile, value);
    }

    public string PrimaryFilePath
    {
        get => _primaryFilePath;
        set => SetProperty(ref _primaryFilePath, value);
    }

    public string PrimaryFilePassword
    {
        get => _primaryFilePassword;
        set => SetProperty(ref _primaryFilePassword, value);
    }

    public string SelectedPrimaryWorksheet
    {
        get => _selectedPrimaryWorksheet;
        set
        {
            if (SetProperty(ref _selectedPrimaryWorksheet, value) && !string.IsNullOrEmpty(value))
                // 工作表变更后，加载工作表信息
                LoadPrimaryWorksheetInfoAsync().ConfigureAwait(false);
        }
    }

    public ObservableCollection<string> PrimaryWorksheets
    {
        get => _primaryWorksheets;
        set => SetProperty(ref _primaryWorksheets, value);
    }

    public ObservableCollection<string> SelectedPrimaryWorksheets
    {
        get => _selectedPrimaryWorksheets;
        set => SetProperty(ref _selectedPrimaryWorksheets, value);
    }

    public ObservableCollection<string> PrimaryColumns
    {
        get => _primaryColumns;
        set
        {
            if (SetProperty(ref _primaryColumns, value)) NotifyCommandsCanExecuteChanged();
        }
    }

    public DataTable? PrimaryPreviewData
    {
        get => _primaryPreviewData;
        set => SetProperty(ref _primaryPreviewData, value);
    }

    // 辅助表文件属性
    public ExcelFile SecondaryFile
    {
        get => _secondaryFile;
        set => SetProperty(ref _secondaryFile, value);
    }

    public string SecondaryFilePath
    {
        get => _secondaryFilePath;
        set => SetProperty(ref _secondaryFilePath, value);
    }

    public string SecondaryFilePassword
    {
        get => _secondaryFilePassword;
        set => SetProperty(ref _secondaryFilePassword, value);
    }

    public string SelectedSecondaryWorksheet
    {
        get => _selectedSecondaryWorksheet;
        set
        {
            if (SetProperty(ref _selectedSecondaryWorksheet, value) && !string.IsNullOrEmpty(value))
                // 工作表变更后，加载工作表信息
                LoadSecondaryWorksheetInfoAsync().ConfigureAwait(false);
        }
    }

    public ObservableCollection<string> SecondaryWorksheets
    {
        get => _secondaryWorksheets;
        set => SetProperty(ref _secondaryWorksheets, value);
    }

    public ObservableCollection<string> SelectedSecondaryWorksheets
    {
        get => _selectedSecondaryWorksheets;
        set => SetProperty(ref _selectedSecondaryWorksheets, value);
    }

    public ObservableCollection<string> SecondaryColumns
    {
        get => _secondaryColumns;
        set
        {
            if (SetProperty(ref _secondaryColumns, value)) NotifyCommandsCanExecuteChanged();
        }
    }

    public DataTable? SecondaryPreviewData
    {
        get => _secondaryPreviewData;
        set => SetProperty(ref _secondaryPreviewData, value);
    }

    // 字段匹配属性
    public ObservableCollection<string> SelectedPrimaryMatchFields
    {
        get => _selectedPrimaryMatchFields;
        set => SetProperty(ref _selectedPrimaryMatchFields, value);
    }

    public ObservableCollection<string> SelectedSecondaryMatchFields
    {
        get => _selectedSecondaryMatchFields;
        set => SetProperty(ref _selectedSecondaryMatchFields, value);
    }

    public ObservableCollection<FieldMapping> FieldMappings
    {
        get => _fieldMappings;
        set => SetProperty(ref _fieldMappings, value);
    }

    // 数据筛选属性
    public ObservableCollection<FilterCondition> PrimaryFilters
    {
        get => _primaryFilters;
        set => SetProperty(ref _primaryFilters, value);
    }

    public ObservableCollection<FilterCondition> SecondaryFilters
    {
        get => _secondaryFilters;
        set => SetProperty(ref _secondaryFilters, value);
    }

    // 进度和状态属性
    public int ProgressValue
    {
        get => _progressValue;
        set => SetProperty(ref _progressValue, value);
    }

    public int ProgressMaximum
    {
        get => _progressMaximum;
        set => SetProperty(ref _progressMaximum, value);
    }

    public string StatusMessage
    {
        get => _statusMessage;
        set => SetProperty(ref _statusMessage, value);
    }

    public bool IsBusy
    {
        get => _isBusy;
        set
        {
            if (SetProperty(ref _isBusy, value)) NotifyCommandsCanExecuteChanged();
        }
    }

    // 过滤器操作符选项
    public IEnumerable<FilterOperator> FilterOperators => Enum.GetValues(typeof(FilterOperator)).Cast<FilterOperator>();

    // 逻辑操作符选项
    public IEnumerable<LogicalOperator> LogicalOperators =>
        Enum.GetValues(typeof(LogicalOperator)).Cast<LogicalOperator>();

    /// <summary>
    ///     是否有主表预览数据
    /// </summary>
    public bool HasPrimaryPreviewData => PrimaryPreviewData != null && PrimaryPreviewData.Rows.Count > 0;

    /// <summary>
    ///     是否有辅助表预览数据
    /// </summary>
    public bool HasSecondaryPreviewData => SecondaryPreviewData != null && SecondaryPreviewData.Rows.Count > 0;

    /// <summary>
    ///     主表文件是否无效
    /// </summary>
    public bool IsPrimaryFileInvalid =>
        !string.IsNullOrEmpty(PrimaryFilePath) && (PrimaryFile == null || !PrimaryFile.IsLoaded);

    /// <summary>
    ///     辅助表文件是否无效
    /// </summary>
    public bool IsSecondaryFileInvalid =>
        !string.IsNullOrEmpty(SecondaryFilePath) && (SecondaryFile == null || !SecondaryFile.IsLoaded);

    /// <summary>
    ///     通知UI相关属性已更改
    /// </summary>
    private void NotifyUIPropertiesChanged()
    {
        OnPropertyChanged(nameof(HasPrimaryPreviewData));
        OnPropertyChanged(nameof(HasSecondaryPreviewData));
        OnPropertyChanged(nameof(IsPrimaryFileInvalid));
        OnPropertyChanged(nameof(IsSecondaryFileInvalid));
    }

    /// <summary>
    ///     重写属性更改通知以包含UI属性
    /// </summary>
    protected override void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        base.OnPropertyChanged(propertyName);

        // 当相关属性变更时，通知UI属性也已变更
        if (propertyName == nameof(PrimaryPreviewData) ||
            propertyName == nameof(SecondaryPreviewData) ||
            propertyName == nameof(PrimaryFile) ||
            propertyName == nameof(SecondaryFile) ||
            propertyName == nameof(PrimaryFilePath) ||
            propertyName == nameof(SecondaryFilePath))
            NotifyUIPropertiesChanged();
    }

    #endregion

    #region 命令

    // 文件浏览命令
    public ICommand BrowsePrimaryFileCommand => _browsePrimaryFileCommand;
    public ICommand BrowseSecondaryFileCommand => _browseSecondaryFileCommand;

    // 加载文件命令
    public ICommand LoadPrimaryFileCommand => _loadPrimaryFileCommand;
    public ICommand LoadSecondaryFileCommand => _loadSecondaryFileCommand;

    // 字段映射命令
    public ICommand AddFieldMappingCommand => _addFieldMappingCommand;
    public ICommand RemoveFieldMappingCommand => _removeFieldMappingCommand;

    // 筛选条件命令
    public ICommand AddPrimaryFilterCommand => _addPrimaryFilterCommand;
    public ICommand RemovePrimaryFilterCommand => _removePrimaryFilterCommand;
    public ICommand AddSecondaryFilterCommand => _addSecondaryFilterCommand;
    public ICommand RemoveSecondaryFilterCommand => _removeSecondaryFilterCommand;

    // 合并和配置命令
    public ICommand StartMergeCommand => _startMergeCommand;
    public ICommand SaveConfigurationCommand => _saveConfigurationCommand;
    public ICommand LoadConfigurationCommand => _loadConfigurationCommand;
    public ICommand ResetConfigurationCommand => _resetConfigurationCommand;

    // ListBox选择更新命令
    public ICommand UpdateSelectedPrimaryMatchFieldsCommand => _updateSelectedPrimaryMatchFieldsCommand;
    public ICommand UpdateSelectedSecondaryMatchFieldsCommand => _updateSelectedSecondaryMatchFieldsCommand;
    public ICommand UpdateSelectedPrimaryWorksheetsCommand => _updateSelectedPrimaryWorksheetsCommand;
    public ICommand UpdateSelectedSecondaryWorksheetsCommand => _updateSelectedSecondaryWorksheetsCommand;

    #endregion

    #region 命令实现

    /// <summary>
    ///     删除配置
    /// </summary>
    private async Task DeleteConfigurationAsync(Configuration? configuration)
    {
        if (configuration == null || string.IsNullOrEmpty(configuration.Path))
            return;

        try
        {
            // 创建MD3风格的确认对话框内容
            var confirmContent = new StackPanel { Margin = new Thickness(24), MinWidth = 400 };

            // 标题区域
            var titlePanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 20)
            };

            // 警告图标
            var warningIcon = new PackIcon
                {
                    Kind = PackIconKind.AlertOutline,
                    Width = 28,
                    Height = 28,
                    Foreground = new SolidColorBrush(Color.FromRgb(255, 152, 0)), // Orange
                    VerticalAlignment = VerticalAlignment.Center
                }
                ;
            titlePanel.Children.Add(warningIcon);

            // 标题文本
            var titleText = new TextBlock
            {
                Text = "确认删除配置",
                FontSize = 20,
                FontWeight = FontWeights.Medium,
                Foreground = new SolidColorBrush(Color.FromRgb(255, 152, 0)),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(12, 0, 0, 0)
            };
            titlePanel.Children.Add(titleText);
            confirmContent.Children.Add(titlePanel);

            // 主要消息
            var messageText = new TextBlock
            {
                Text = $"您确定要删除配置 \"{configuration.Name}\" 吗？",
                FontSize = 16,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 12)
            };
            confirmContent.Children.Add(messageText);

            // 警告信息
            var warningText = new TextBlock
            {
                Text = "此操作无法撤销。",
                FontWeight = FontWeights.Medium,
                Foreground = new SolidColorBrush(Color.FromRgb(244, 67, 54)), // Red
                Margin = new Thickness(0, 0, 0, 32)
            };
            confirmContent.Children.Add(warningText);

            // 按钮区域
            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            var cancelButton = new Button
            {
                Content = "取消",
                Style = Application.Current.Resources["MD3OutlinedButton"] as Style,
                Margin = new Thickness(0, 0, 12, 0),
                MinWidth = 88
            };

            var deleteButton = new Button
            {
                Content = "删除",
                Style = Application.Current.Resources["MD3FilledButton"] as Style,
                Background = new SolidColorBrush(Color.FromRgb(244, 67, 54)), // Red
                MinWidth = 88
            };

            buttonPanel.Children.Add(cancelButton);
            buttonPanel.Children.Add(deleteButton);
            confirmContent.Children.Add(buttonPanel);

            // 显示确认对话框并等待结果
            var dialogResult = false;

            cancelButton.Click += (s, e) => { DialogHost.Close("ConfirmDialog"); };
            deleteButton.Click += (s, e) =>
            {
                dialogResult = true;
                DialogHost.Close("ConfirmDialog");
            };

            // 使用不同的DialogHost标识符
            await DialogHost.Show(confirmContent, "ConfirmDialog");

            if (!dialogResult)
                return;

            // 执行删除操作
            IsBusy = true;
            StatusMessage = $"正在删除配置 \"{configuration.Name}\"...";

            var success = await _configurationManager.DeleteConfigurationAsync(configuration.Path);

            if (success)
            {
                // 创建成功消息对话框
                var successContent = new StackPanel { Margin = new Thickness(24), MinWidth = 350 };

                // 成功标题区域
                var successTitlePanel = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Margin = new Thickness(0, 0, 0, 20)
                };

                var successIcon = new PackIcon
                    {
                        Kind = PackIconKind.CheckCircleOutline,
                        Width = 28,
                        Height = 28,
                        Foreground = new SolidColorBrush(Color.FromRgb(76, 175, 80)), // Green
                        VerticalAlignment = VerticalAlignment.Center
                    }
                    ;
                successTitlePanel.Children.Add(successIcon);

                var successTitleText = new TextBlock
                {
                    Text = "删除成功",
                    FontSize = 20,
                    FontWeight = FontWeights.Medium,
                    Foreground = new SolidColorBrush(Color.FromRgb(76, 175, 80)),
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(12, 0, 0, 0)
                };
                successTitlePanel.Children.Add(successTitleText);
                successContent.Children.Add(successTitlePanel);

                var successMessageText = new TextBlock
                {
                    Text = $"配置 \"{configuration.Name}\" 已成功删除。",
                    FontSize = 14,
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 0, 0, 32)
                };
                successContent.Children.Add(successMessageText);

                var successOkButton = new Button
                {
                    Content = "确定",
                    Style = Application.Current.Resources["MD3FilledButton"] as Style,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    MinWidth = 88,
                    IsDefault = true
                };

                successOkButton.Click += (s, e) => { DialogHost.Close("ConfirmDialog"); };
                successContent.Children.Add(successOkButton);

                await DialogHost.Show(successContent, "ConfirmDialog");

                StatusMessage = "配置删除完成";
            }
            else
            {
                throw new InvalidOperationException("配置文件不存在或已被删除");
            }
        }
        catch (Exception ex)
        {
            // 创建错误消息对话框
            var errorContent = new StackPanel { Margin = new Thickness(24), MinWidth = 400 };

            // 错误标题区域
            var errorTitlePanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 20)
            };

            var errorIcon = new PackIcon
                {
                    Kind = PackIconKind.AlertCircleOutline,
                    Width = 28,
                    Height = 28,
                    Foreground = new SolidColorBrush(Color.FromRgb(244, 67, 54)), // Red
                    VerticalAlignment = VerticalAlignment.Center
                }
                ;
            errorTitlePanel.Children.Add(errorIcon);

            var errorTitleText = new TextBlock
            {
                Text = "删除失败",
                FontSize = 20,
                FontWeight = FontWeights.Medium,
                Foreground = new SolidColorBrush(Color.FromRgb(244, 67, 54)),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(12, 0, 0, 0)
            };
            errorTitlePanel.Children.Add(errorTitleText);
            errorContent.Children.Add(errorTitlePanel);

            var errorMessageText = new TextBlock
            {
                Text = $"删除配置时出错: {ex.Message}",
                FontSize = 14,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 32)
            };
            errorContent.Children.Add(errorMessageText);

            var errorOkButton = new Button
            {
                Content = "确定",
                Style = Application.Current.Resources["MD3OutlinedButton"] as Style,
                HorizontalAlignment = HorizontalAlignment.Right,
                MinWidth = 88,
                IsDefault = true
            };

            errorOkButton.Click += (s, e) => { DialogHost.Close("ConfirmDialog"); };
            errorContent.Children.Add(errorOkButton);

            await DialogHost.Show(errorContent, "ConfirmDialog");

            StatusMessage = "删除配置失败";
            Debug.WriteLine($"删除配置异常: {ex}");
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>
    ///     打开配置文件目录
    /// </summary>
    private void OpenConfigurationDirectory()
    {
        try
        {
            _configurationManager.OpenConfigurationDirectory();
            StatusMessage = "已打开配置文件目录";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"打开配置目录时出错: {ex.Message}", "错误",
                MessageBoxButton.OK, MessageBoxImage.Error);
            StatusMessage = "打开配置目录失败";
            Debug.WriteLine($"打开配置目录异常: {ex}");
        }
    }

    // 通知所有命令重新检查可执行状态
    private void NotifyCommandsCanExecuteChanged()
    {
        _addFieldMappingCommand?.NotifyCanExecuteChanged();
        _addPrimaryFilterCommand?.NotifyCanExecuteChanged();
        _addSecondaryFilterCommand?.NotifyCanExecuteChanged();
        _startMergeCommand?.NotifyCanExecuteChanged();
    }

    // 处理筛选条件变更，自动刷新预览
    private void FilterConditionsCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        // 当筛选条件集合发生变化时，刷新预览数据
        if (e.Action != NotifyCollectionChangedAction.Move) LoadPreviewDataWithFiltersAsync().ConfigureAwait(false);
    }

    // 监听筛选条件属性变更
    private void FilterCondition_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        // 当筛选条件属性变更时，刷新预览数据
        LoadPreviewDataWithFiltersAsync().ConfigureAwait(false);
    }

    // 更新主表选中字段
    private void UpdateSelectedPrimaryMatchFields(IList items)
    {
        if (items == null) return;

        SelectedPrimaryMatchFields.Clear();
        foreach (var item in items) SelectedPrimaryMatchFields.Add(item.ToString());
        NotifyCommandsCanExecuteChanged();
    }

    // 更新辅助表选中字段
    private void UpdateSelectedSecondaryMatchFields(IList items)
    {
        if (items == null) return;

        SelectedSecondaryMatchFields.Clear();
        foreach (var item in items) SelectedSecondaryMatchFields.Add(item.ToString());
        NotifyCommandsCanExecuteChanged();
    }

    // 更新主表选中工作表
    private void UpdateSelectedPrimaryWorksheets(IList items)
    {
        if (items == null) return;

        SelectedPrimaryWorksheets.Clear();
        foreach (var item in items) SelectedPrimaryWorksheets.Add(item.ToString());

        // 更新主表工作表信息
        if (SelectedPrimaryWorksheets.Count > 0) LoadPrimaryWorksheetsInfoAsync().ConfigureAwait(false);
    }

    // 更新辅助表选中工作表
    private void UpdateSelectedSecondaryWorksheets(IList items)
    {
        if (items == null) return;

        SelectedSecondaryWorksheets.Clear();
        foreach (var item in items) SelectedSecondaryWorksheets.Add(item.ToString());

        // 更新辅助表工作表信息
        if (SelectedSecondaryWorksheets.Count > 0) LoadSecondaryWorksheetsInfoAsync().ConfigureAwait(false);
    }

    // 浏览主表文件
    private async Task BrowsePrimaryFileAsync()
    {
        var openFileDialog = new OpenFileDialog
        {
            Filter = "Excel文件|*.xlsx;*.xls|所有文件|*.*",
            Title = "选择主表Excel文件"
        };

        if (openFileDialog.ShowDialog() == true)
        {
            PrimaryFilePath = openFileDialog.FileName;
            await LoadPrimaryFileAsync();
        }
    }

    // 浏览辅助表文件
    private async Task BrowseSecondaryFileAsync()
    {
        var openFileDialog = new OpenFileDialog
        {
            Filter = "Excel文件|*.xlsx;*.xls|所有文件|*.*",
            Title = "选择辅助表Excel文件"
        };

        if (openFileDialog.ShowDialog() == true)
        {
            SecondaryFilePath = openFileDialog.FileName;
            await LoadSecondaryFileAsync();
        }
    }

    /// <summary>
    ///     加载主表文件
    /// </summary>
    private async Task LoadPrimaryFileAsync(bool isRefresh = false)
    {
        if (string.IsNullOrEmpty(PrimaryFilePath))
            return;

        try
        {
            IsBusy = true;
            StatusMessage = "正在加载主表文件...";

            // 如果是刷新操作，先完全关闭之前的文件
            if (isRefresh && PrimaryFile != null && PrimaryFile.IsLoaded)
            {
                await _excelFileManager.CloseFileAsync(PrimaryFile);
                PrimaryFile = null;

                // 执行垃圾回收以确保资源释放
                GC.Collect();
                GC.WaitForPendingFinalizers();

                // 短暂延时以确保文件句柄完全释放
                await Task.Delay(100);
            }

            // 重置相关数据
            PrimaryWorksheets.Clear();
            PrimaryColumns.Clear();

            // 如果是刷新操作，不清除已选工作表和匹配字段
            if (!isRefresh)
            {
                SelectedPrimaryWorksheets.Clear();
                SelectedPrimaryMatchFields.Clear();
                PrimaryPreviewData = null;
            }

            // 加载文件
            PrimaryFile = await _excelFileManager.LoadExcelFileAsync(PrimaryFilePath, PrimaryFilePassword);

            // 更新UI
            foreach (var worksheet in PrimaryFile.Worksheets) PrimaryWorksheets.Add(worksheet);

            if (PrimaryWorksheets.Count > 0 && (!isRefresh || SelectedPrimaryWorksheets.Count == 0))
            {
                SelectedPrimaryWorksheets.Add(PrimaryWorksheets[0]);
                SelectedPrimaryWorksheet = PrimaryWorksheets[0];
            }

            StatusMessage = $"主表文件已加载，共{PrimaryFile.WorksheetCount}个工作表";

            CheckFileSize();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"加载主表文件时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            StatusMessage = "加载主表文件失败";
            Debug.WriteLine($"加载主表文件异常: {ex}");
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>
    ///     加载辅助表文件
    /// </summary>
    private async Task LoadSecondaryFileAsync(bool isRefresh = false)
    {
        if (string.IsNullOrEmpty(SecondaryFilePath))
            return;

        try
        {
            IsBusy = true;
            StatusMessage = "正在加载辅助表文件...";

            // 如果是刷新操作，先完全关闭之前的文件
            if (isRefresh && SecondaryFile != null && SecondaryFile.IsLoaded)
            {
                await _excelFileManager.CloseFileAsync(SecondaryFile);
                SecondaryFile = null;

                // 执行垃圾回收以确保资源释放
                GC.Collect();
                GC.WaitForPendingFinalizers();

                // 短暂延时以确保文件句柄完全释放
                await Task.Delay(100);
            }

            // 重置相关数据
            SecondaryWorksheets.Clear();
            SecondaryColumns.Clear();

            // 如果是刷新操作，不清除已选工作表和匹配字段
            if (!isRefresh)
            {
                SelectedSecondaryWorksheets.Clear();
                SelectedSecondaryMatchFields.Clear();
                SecondaryPreviewData = null;
            }

            // 加载文件
            SecondaryFile = await _excelFileManager.LoadExcelFileAsync(SecondaryFilePath, SecondaryFilePassword);

            // 更新UI
            foreach (var worksheet in SecondaryFile.Worksheets) SecondaryWorksheets.Add(worksheet);

            if (SecondaryWorksheets.Count > 0 && (!isRefresh || SelectedSecondaryWorksheets.Count == 0))
            {
                SelectedSecondaryWorksheets.Add(SecondaryWorksheets[0]);
                SelectedSecondaryWorksheet = SecondaryWorksheets[0];
            }

            StatusMessage = $"辅助表文件已加载，共{SecondaryFile.WorksheetCount}个工作表";

            CheckFileSize();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"加载辅助表文件时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            StatusMessage = "加载辅助表文件失败";
            Debug.WriteLine($"加载辅助表文件异常: {ex}");
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>
    ///     检查文件大小并建议是否关闭预览
    /// </summary>
    private async void CheckFileSize()
    {
        try
        {
            var primarySize = !string.IsNullOrEmpty(PrimaryFilePath) && File.Exists(PrimaryFilePath)
                ? new FileInfo(PrimaryFilePath).Length
                : 0;
            var secondarySize = !string.IsNullOrEmpty(SecondaryFilePath) && File.Exists(SecondaryFilePath)
                ? new FileInfo(SecondaryFilePath).Length
                : 0;

            // 如果任一文件超过阈值，建议关闭预览
            const long largeFileThreshold = 10 * 1024 * 1024; // 10MB
            const long veryLargeFileThreshold = 50 * 1024 * 1024; // 50MB

            var maxSize = Math.Max(primarySize, secondarySize);
            var totalSize = primarySize + secondarySize;

            // 检查是否需要建议关闭预览
            if (IsPreviewEnabled && (maxSize > largeFileThreshold || totalSize > veryLargeFileThreshold))
            {
                var fileSizeInfo = new
                {
                    PrimarySize = primarySize,
                    SecondarySize = secondarySize,
                    MaxSize = maxSize,
                    TotalSize = totalSize,
                    IsVeryLarge = maxSize > veryLargeFileThreshold
                };

                await ShowSuggestDisablePreviewDialog(fileSizeInfo);
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"检查文件大小时出错: {ex.Message}");
            // 不影响主流程，仅记录错误
        }
    }

    /// <summary>
    ///     格式化文件大小显示
    /// </summary>
    private string FormatFileSize(long bytes)
    {
        if (bytes >= 1024 * 1024 * 1024)
            return $"{bytes / (1024.0 * 1024 * 1024):F1} GB";
        if (bytes >= 1024 * 1024)
            return $"{bytes / (1024.0 * 1024):F1} MB";
        if (bytes >= 1024)
            return $"{bytes / 1024.0:F1} KB";
        return $"{bytes} B";
    }

    /// <summary>
    ///     显示建议关闭预览的对话框
    /// </summary>
    private async Task ShowSuggestDisablePreviewDialog(dynamic fileSizeInfo)
    {
        try
        {
            // 创建MD3风格的建议对话框内容
            var suggestionContent = new StackPanel { Margin = new Thickness(24), MinWidth = 500 };

            // 标题区域
            var titlePanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 20)
            };

            // 性能警告图标
            var warningIcon = new PackIcon
            {
                Kind = PackIconKind.Speedometer,
                Width = 32,
                Height = 32,
                Foreground = new SolidColorBrush(Color.FromRgb(255, 193, 7)), // Amber
                VerticalAlignment = VerticalAlignment.Center
            };
            titlePanel.Children.Add(warningIcon);

            // 标题文本
            var titleText = new TextBlock
            {
                Text = "检测到大文件",
                FontSize = 20,
                FontWeight = FontWeights.Medium,
                Foreground = new SolidColorBrush(Color.FromRgb(255, 193, 7)),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(12, 0, 0, 0)
            };
            titlePanel.Children.Add(titleText);
            suggestionContent.Children.Add(titlePanel);

            // 文件大小信息卡片
            var infoCard = new Card
            {
                Style = Application.Current.Resources["MD3OutlinedCard"] as Style,
                Background = new SolidColorBrush(Color.FromRgb(255, 248, 225)), // Light amber
                Margin = new Thickness(0, 0, 0, 20)
            };

            var infoPanel = new StackPanel();

            // 文件大小标题
            var infoTitle = new TextBlock
            {
                Text = "文件大小信息",
                FontWeight = FontWeights.Medium,
                Margin = new Thickness(0, 0, 0, 12),
                Style = Application.Current.Resources["MD3TitleMedium"] as Style
            };
            infoPanel.Children.Add(infoTitle);

            // 主表文件大小
            if (fileSizeInfo.PrimarySize > 0)
            {
                var primarySizeText = new TextBlock
                {
                    Margin = new Thickness(0, 0, 0, 8),
                    Style = Application.Current.Resources["MD3BodyMedium"] as Style
                };
                primarySizeText.Inlines.Add(new Run { Text = "主表文件: ", FontWeight = FontWeights.Medium });
                primarySizeText.Inlines.Add(new Run { Text = FormatFileSize(fileSizeInfo.PrimarySize) });
                infoPanel.Children.Add(primarySizeText);
            }

            // 辅助表文件大小
            if (fileSizeInfo.SecondarySize > 0)
            {
                var secondarySizeText = new TextBlock
                {
                    Margin = new Thickness(0, 0, 0, 8),
                    Style = Application.Current.Resources["MD3BodyMedium"] as Style
                };
                secondarySizeText.Inlines.Add(new Run { Text = "辅助表文件: ", FontWeight = FontWeights.Medium });
                secondarySizeText.Inlines.Add(new Run { Text = FormatFileSize(fileSizeInfo.SecondarySize) });
                infoPanel.Children.Add(secondarySizeText);
            }

            // 总大小
            var totalSizeText = new TextBlock
            {
                Style = Application.Current.Resources["MD3BodyMedium"] as Style
            };
            totalSizeText.Inlines.Add(new Run { Text = "总大小: ", FontWeight = FontWeights.Medium });
            totalSizeText.Inlines.Add(new Run { Text = FormatFileSize(fileSizeInfo.TotalSize) });
            infoPanel.Children.Add(totalSizeText);

            infoCard.Content = infoPanel;
            suggestionContent.Children.Add(infoCard);

            // 主要消息
            var messageText = new TextBlock
            {
                Text = fileSizeInfo.IsVeryLarge
                    ? "检测到超大文件，强烈建议关闭数据预览以优化性能。"
                    : "检测到较大文件，建议关闭数据预览以提高加载速度。",
                FontSize = 16,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 16)
            };
            suggestionContent.Children.Add(messageText);

            // 说明信息
            var explanationText = new TextBlock
            {
                Text = "关闭数据预览后：\n• 文件加载速度显著提升\n• 内存占用大幅减少\n• 所有数据处理功能正常使用\n• 可随时重新启用预览功能",
                Style = Application.Current.Resources["MD3BodyMedium"] as Style,
                Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)), // Gray
                Margin = new Thickness(0, 0, 0, 32),
                LineHeight = 20
            };
            suggestionContent.Children.Add(explanationText);

            // 记住选择的复选框
            var rememberCheckBox = new CheckBox
            {
                Content = "记住我的选择（下次不再提示）",
                Style = Application.Current.Resources["MaterialDesignCheckBox"] as Style,
                Margin = new Thickness(0, 0, 0, 24),
                IsChecked = false
            };
            suggestionContent.Children.Add(rememberCheckBox);

            // 按钮区域
            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            var keepPreviewButton = new Button
            {
                Content = "保持预览开启",
                Style = Application.Current.Resources["MD3OutlinedButton"] as Style,
                Margin = new Thickness(0, 0, 12, 0),
                MinWidth = 120
            };

            var disablePreviewButton = new Button
            {
                Content = fileSizeInfo.IsVeryLarge ? "关闭预览（推荐）" : "关闭预览",
                Style = Application.Current.Resources["MD3FilledButton"] as Style,
                Background = new SolidColorBrush(Color.FromRgb(255, 193, 7)), // Amber
                MinWidth = 120
            };

            buttonPanel.Children.Add(keepPreviewButton);
            buttonPanel.Children.Add(disablePreviewButton);
            suggestionContent.Children.Add(buttonPanel);

            // 显示对话框并等待结果
            var dialogResult = false;

            keepPreviewButton.Click += (s, e) => { DialogHost.Close("RootDialog"); };
            disablePreviewButton.Click += (s, e) =>
            {
                dialogResult = true;
                DialogHost.Close("RootDialog");
            };

            // 使用DialogHost显示对话框
            await DialogHost.Show(suggestionContent, "RootDialog");

            // 处理用户选择
            if (dialogResult)
            {
                // 用户选择关闭预览
                IsPreviewEnabled = false;
                StatusMessage = "已关闭数据预览以优化性能";

                // 如果用户选择记住选择，可以保存到配置中
                if (rememberCheckBox.IsChecked == true)
                    // TODO: 可以添加到用户配置中，下次自动应用
                    Debug.WriteLine("用户选择记住关闭预览的选择");
            }
            else
            {
                // 用户选择保持预览
                StatusMessage = "保持数据预览开启";
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"显示文件大小建议对话框时出错: {ex.Message}");
            // 如果对话框显示失败，静默处理，不影响主流程
        }
    }

    // 加载主表工作表信息
    private async Task LoadPrimaryWorksheetInfoAsync()
    {
        if (PrimaryFile == null || string.IsNullOrEmpty(SelectedPrimaryWorksheet))
            return;

        try
        {
            IsBusy = true;
            StatusMessage = $"正在加载主表工作表 {SelectedPrimaryWorksheet} 信息...";

            // 更新已选工作表
            PrimaryFile.SelectedWorksheet = SelectedPrimaryWorksheet;

            // 加载工作表信息
            PrimaryFile = await _excelFileManager.LoadWorksheetInfoAsync(PrimaryFile);

            // 更新UI
            PrimaryColumns.Clear();
            foreach (var column in PrimaryFile.Columns) PrimaryColumns.Add(column);

            // 只有在预览启用时才加载预览数据
            if (IsPreviewEnabled) await LoadPreviewDataWithFiltersAsync();

            StatusMessage = $"主表工作表已加载，共{PrimaryFile.RowCount}行，{PrimaryFile.ColumnCount}列";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"加载主表工作表信息时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            StatusMessage = "加载主表工作表信息失败";
            Debug.WriteLine($"加载主表工作表信息异常: {ex}");
        }
        finally
        {
            IsBusy = false;
            NotifyCommandsCanExecuteChanged(); // 确保在完成后更新命令状态
        }
    }

    // 加载多个主表工作表信息
    private async Task LoadPrimaryWorksheetsInfoAsync()
    {
        if (PrimaryFile == null || SelectedPrimaryWorksheets.Count == 0)
            return;

        try
        {
            IsBusy = true;
            StatusMessage = "正在加载主表工作表信息...";

            // 更新已选工作表
            PrimaryFile.SelectedWorksheets.Clear();
            foreach (var worksheet in SelectedPrimaryWorksheets) PrimaryFile.SelectedWorksheets.Add(worksheet);

            // 如果只有一个工作表，保持向后兼容
            if (SelectedPrimaryWorksheets.Count == 1) PrimaryFile.SelectedWorksheet = SelectedPrimaryWorksheets[0];

            // 加载每个工作表的信息
            foreach (var worksheetName in SelectedPrimaryWorksheets)
            {
                // 临时设置当前工作表
                PrimaryFile.SelectedWorksheet = worksheetName;

                // 加载工作表信息
                await _excelFileManager.LoadWorksheetInfoAsync(PrimaryFile);

                // 存储工作表列信息
                if (!PrimaryFile.WorksheetInfo.ContainsKey(worksheetName))
                    PrimaryFile.WorksheetInfo[worksheetName] = (
                        PrimaryFile.RowCount,
                        PrimaryFile.ColumnCount,
                        new List<string>(PrimaryFile.Columns)
                    );
            }

            // 合并所有工作表的列
            PrimaryColumns.Clear();
            var uniqueColumns = new HashSet<string>();
            foreach (var (_, _, columns) in PrimaryFile.WorksheetInfo.Values)
            foreach (var column in columns)
                uniqueColumns.Add(column);

            foreach (var column in uniqueColumns) PrimaryColumns.Add(column);

            // 更新数据预览
            await LoadPreviewDataWithFiltersAsync();

            var totalRows = PrimaryFile.WorksheetInfo.Values.Sum(info => info.RowCount);
            StatusMessage = $"主表工作表已加载，共{totalRows}行，{PrimaryColumns.Count}列";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"加载主表工作表信息时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            StatusMessage = "加载主表工作表信息失败";
        }
        finally
        {
            IsBusy = false;
            NotifyCommandsCanExecuteChanged();
        }
    }

    // 加载辅助表工作表信息
    private async Task LoadSecondaryWorksheetInfoAsync()
    {
        if (SecondaryFile == null || string.IsNullOrEmpty(SelectedSecondaryWorksheet))
            return;

        try
        {
            IsBusy = true;
            StatusMessage = $"正在加载辅助表工作表 {SelectedSecondaryWorksheet} 信息...";

            // 更新已选工作表
            SecondaryFile.SelectedWorksheet = SelectedSecondaryWorksheet;

            // 加载工作表信息
            SecondaryFile = await _excelFileManager.LoadWorksheetInfoAsync(SecondaryFile);

            // 更新UI
            SecondaryColumns.Clear();
            foreach (var column in SecondaryFile.Columns) SecondaryColumns.Add(column);

            // 加载预览数据
            await LoadPreviewDataWithFiltersAsync();

            StatusMessage = $"辅助表工作表已加载，共{SecondaryFile.RowCount}行，{SecondaryFile.ColumnCount}列";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"加载辅助表工作表信息时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            StatusMessage = "加载辅助表工作表信息失败";
            Debug.WriteLine($"加载辅助表工作表信息异常: {ex}");
        }
        finally
        {
            IsBusy = false;
            NotifyCommandsCanExecuteChanged(); // 确保在完成后更新命令状态
        }
    }

    // 加载多个辅助表工作表信息
    private async Task LoadSecondaryWorksheetsInfoAsync()
    {
        if (SecondaryFile == null || SelectedSecondaryWorksheets.Count == 0)
            return;

        try
        {
            IsBusy = true;
            StatusMessage = "正在加载辅助表工作表信息...";

            // 更新已选工作表
            SecondaryFile.SelectedWorksheets.Clear();
            foreach (var worksheet in SelectedSecondaryWorksheets) SecondaryFile.SelectedWorksheets.Add(worksheet);

            // 如果只有一个工作表，保持向后兼容
            if (SelectedSecondaryWorksheets.Count == 1)
                SecondaryFile.SelectedWorksheet = SelectedSecondaryWorksheets[0];

            // 加载每个工作表的信息
            foreach (var worksheetName in SelectedSecondaryWorksheets)
            {
                // 临时设置当前工作表
                SecondaryFile.SelectedWorksheet = worksheetName;

                // 加载工作表信息
                await _excelFileManager.LoadWorksheetInfoAsync(SecondaryFile);

                // 存储工作表列信息
                if (!SecondaryFile.WorksheetInfo.ContainsKey(worksheetName))
                    SecondaryFile.WorksheetInfo[worksheetName] = (
                        SecondaryFile.RowCount,
                        SecondaryFile.ColumnCount,
                        new List<string>(SecondaryFile.Columns)
                    );
            }

            // 合并所有工作表的列
            SecondaryColumns.Clear();
            var uniqueColumns = new HashSet<string>();
            foreach (var (_, _, columns) in SecondaryFile.WorksheetInfo.Values)
            foreach (var column in columns)
                uniqueColumns.Add(column);

            foreach (var column in uniqueColumns) SecondaryColumns.Add(column);

            // 更新数据预览
            await LoadPreviewDataWithFiltersAsync();

            var totalRows = SecondaryFile.WorksheetInfo.Values.Sum(info => info.RowCount);
            StatusMessage = $"辅助表工作表已加载，共{totalRows}行，{SecondaryColumns.Count}列";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"加载辅助表工作表信息时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            StatusMessage = "加载辅助表工作表信息失败";
        }
        finally
        {
            IsBusy = false;
            NotifyCommandsCanExecuteChanged();
        }
    }

    // 加载预览数据并应用筛选条件
    private async Task LoadPreviewDataWithFiltersAsync()
    {
        // 如果预览被禁用，直接返回
        if (!IsPreviewEnabled) return;
        try
        {
            if (PrimaryFile == null || SecondaryFile == null ||
                string.IsNullOrEmpty(PrimaryFile.SelectedWorksheet) ||
                string.IsNullOrEmpty(SecondaryFile.SelectedWorksheet))
                return;

            IsBusy = true;
            StatusMessage = "正在加载预览数据...";

            // 获取原始数据
            var primaryDataRaw = await _excelFileManager.GetWorksheetDataAsync(PrimaryFile);
            var secondaryDataRaw = await _excelFileManager.GetWorksheetDataAsync(SecondaryFile);

            // 应用筛选条件
            PrimaryPreviewData = _excelFileManager.ApplyFilters(primaryDataRaw, PrimaryFilters.ToList());
            SecondaryPreviewData = _excelFileManager.ApplyFilters(secondaryDataRaw, SecondaryFilters.ToList());

            // 显示结果
            StatusMessage = $"预览数据已加载，主表:{PrimaryPreviewData.Rows.Count}行，辅助表:{SecondaryPreviewData.Rows.Count}行";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"加载预览数据时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            StatusMessage = "加载预览数据失败";
        }
        finally
        {
            IsBusy = false;
        }
    }

    // 添加字段映射
    private void AddFieldMapping()
    {
        try
        {
            // 创建新的字段映射
            var mapping = new FieldMapping
            {
                SourceField = SecondaryColumns.FirstOrDefault() ?? string.Empty,
                TargetField = string.Empty // 可以为空，表示新增字段
            };

            // 监听属性变更以自动更新UI
            ((INotifyPropertyChanged)mapping).PropertyChanged += FieldMapping_PropertyChanged;

            // 添加到映射列表
            FieldMappings.Add(mapping);
            Debug.WriteLine($"添加字段映射: {mapping.SourceField} -> {mapping.TargetField}");

            // 确保更新命令状态
            NotifyCommandsCanExecuteChanged();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"添加字段映射时出错: {ex.Message}");
            MessageBox.Show($"添加字段映射时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // 处理字段映射属性变更
    private void FieldMapping_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        // 通知命令可能需要更新
        NotifyCommandsCanExecuteChanged();
    }

    // 判断是否可以添加字段映射
    private bool CanAddFieldMapping()
    {
        var canAdd = SecondaryColumns != null && SecondaryColumns.Count > 0 && !IsBusy;
        Debug.WriteLine($"CanAddFieldMapping: {canAdd}, SecondaryColumns数量: {SecondaryColumns?.Count ?? 0}");
        return canAdd;
    }

    /// <summary>
    ///     刷新数据
    /// </summary>
    private async Task RefreshDataAsync()
    {
        try
        {
            IsBusy = true;
            StatusMessage = "正在刷新数据...";

            // 记录当前选择的工作表
            var selectedPrimarySheets = new List<string>(SelectedPrimaryWorksheets);
            var selectedSecondarySheets = new List<string>(SelectedSecondaryWorksheets);

            // 完全清除现有数据
            PrimaryWorksheets.Clear();
            SecondaryWorksheets.Clear();
            PrimaryColumns.Clear();
            SecondaryColumns.Clear();
            SelectedPrimaryMatchFields.Clear();
            SelectedSecondaryMatchFields.Clear();
            PrimaryPreviewData = null;
            SecondaryPreviewData = null;

            // 强制关闭文件并释放资源
            if (PrimaryFile != null)
            {
                await _excelFileManager.CloseFileAsync(PrimaryFile);
                PrimaryFile = null;
            }

            if (SecondaryFile != null)
            {
                await _excelFileManager.CloseFileAsync(SecondaryFile);
                SecondaryFile = null;
            }

            // 执行垃圾回收以确保资源释放
            GC.Collect();
            GC.WaitForPendingFinalizers();

            // 等待一小段时间确保文件句柄完全释放
            await Task.Delay(500);

            // 重新加载文件
            if (!string.IsNullOrEmpty(PrimaryFilePath) && File.Exists(PrimaryFilePath))
            {
                // 更新文件的最后修改时间
                if (File.Exists(PrimaryFilePath))
                {
                    var fileInfo = new FileInfo(PrimaryFilePath);
                    fileInfo.Refresh(); // 刷新文件信息
                    StatusMessage = $"正在加载主表文件 (最后修改: {fileInfo.LastWriteTime})...";
                }

                await LoadPrimaryFileAsync(true);
            }

            if (!string.IsNullOrEmpty(SecondaryFilePath) && File.Exists(SecondaryFilePath))
            {
                // 更新文件的最后修改时间
                if (File.Exists(SecondaryFilePath))
                {
                    var fileInfo = new FileInfo(SecondaryFilePath);
                    fileInfo.Refresh(); // 刷新文件信息
                    StatusMessage = $"正在加载辅助表文件 (最后修改: {fileInfo.LastWriteTime})...";
                }

                await LoadSecondaryFileAsync(true);
            }

            // 恢复选择的工作表
            if (PrimaryWorksheets.Count > 0)
            {
                foreach (var sheet in selectedPrimarySheets)
                    if (PrimaryWorksheets.Contains(sheet))
                        SelectedPrimaryWorksheets.Add(sheet);

                // 如果没有工作表被选中，选择第一个
                if (SelectedPrimaryWorksheets.Count == 0 && PrimaryWorksheets.Count > 0)
                    SelectedPrimaryWorksheets.Add(PrimaryWorksheets[0]);

                await LoadPrimaryWorksheetsInfoAsync();
            }

            if (SecondaryWorksheets.Count > 0)
            {
                foreach (var sheet in selectedSecondarySheets)
                    if (SecondaryWorksheets.Contains(sheet))
                        SelectedSecondaryWorksheets.Add(sheet);

                // 如果没有工作表被选中，选择第一个
                if (SelectedSecondaryWorksheets.Count == 0 && SecondaryWorksheets.Count > 0)
                    SelectedSecondaryWorksheets.Add(SecondaryWorksheets[0]);

                await LoadSecondaryWorksheetsInfoAsync();
            }

            // 验证和更新字段映射
            ValidateAndUpdateFieldMappings();

            // 刷新预览数据
            await LoadPreviewDataWithFiltersAsync();

            StatusMessage = "数据已刷新";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"刷新数据时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            StatusMessage = "刷新数据失败";
            Debug.WriteLine($"刷新数据异常: {ex}");
        }
        finally
        {
            IsBusy = false;
            NotifyCommandsCanExecuteChanged();
        }
    }

    /// <summary>
    ///     判断是否可以刷新数据
    /// </summary>
    private bool CanRefreshData()
    {
        return (!string.IsNullOrEmpty(PrimaryFilePath) || !string.IsNullOrEmpty(SecondaryFilePath)) && !IsBusy;
    }

    /// <summary>
    ///     诊断匹配问题
    /// </summary>
    private async Task DiagnoseMatchingAsync()
    {
        if (!ValidateMatchingParameters())
            return;

        try
        {
            IsBusy = true;
            StatusMessage = "正在诊断匹配问题...";

            var result = await _excelFileManager.DiagnoseMatchFieldsAsync(
                PrimaryFile,
                SecondaryFile,
                SelectedPrimaryMatchFields.ToList(),
                SelectedSecondaryMatchFields.ToList());

            // 显示诊断结果对话框
            var diagContent = new StackPanel { Margin = new Thickness(24) };
            diagContent.Children.Add(new TextBlock
            {
                Text = "匹配问题诊断",
                FontSize = 18,
                FontWeight = FontWeights.Medium,
                Margin = new Thickness(0, 0, 0, 16)
            });

            var scrollViewer = new ScrollViewer
            {
                MaxHeight = 400,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto
            };

            var textBox = new TextBox
            {
                Text = result,
                IsReadOnly = true,
                TextWrapping = TextWrapping.NoWrap,
                FontFamily = new FontFamily("Consolas"),
                VerticalAlignment = VerticalAlignment.Stretch,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto
            };

            scrollViewer.Content = textBox;
            diagContent.Children.Add(scrollViewer);

            var copyButton = new Button
            {
                Content = "复制诊断信息",
                Style = Application.Current.Resources["MaterialDesignOutlinedButton"] as Style,
                Margin = new Thickness(0, 16, 8, 0),
                HorizontalAlignment = HorizontalAlignment.Right
            };

            var closeButton = new Button
            {
                Content = "关闭",
                Style = Application.Current.Resources["MaterialDesignOutlinedLightButton"] as Style,
                Margin = new Thickness(0, 16, 0, 0),
                HorizontalAlignment = HorizontalAlignment.Right,
                IsDefault = true
            };

            copyButton.Click += (s, e) =>
            {
                try
                {
                    Clipboard.SetText(result);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"复制到剪贴板时出错: {ex.Message}");
                }
            };

            closeButton.Click += (s, e) => { DialogHost.Close("RootDialog"); };

            var buttonPanel = new StackPanel
                { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            buttonPanel.Children.Add(copyButton);
            buttonPanel.Children.Add(closeButton);
            diagContent.Children.Add(buttonPanel);

            await DialogHost.Show(diagContent, "RootDialog");

            StatusMessage = "诊断完成";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"诊断匹配问题时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            StatusMessage = "诊断匹配问题失败";
            Debug.WriteLine($"诊断匹配问题异常: {ex}");
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>
    ///     判断是否可以诊断匹配
    /// </summary>
    private bool CanDiagnoseMatching()
    {
        return PrimaryFile != null && PrimaryFile.IsLoaded &&
               SecondaryFile != null && SecondaryFile.IsLoaded &&
               SelectedPrimaryMatchFields.Count > 0 &&
               SelectedSecondaryMatchFields.Count > 0 &&
               SelectedPrimaryMatchFields.Count == SelectedSecondaryMatchFields.Count &&
               !IsBusy;
    }

    /// <summary>
    ///     验证匹配参数
    /// </summary>
    private bool ValidateMatchingParameters()
    {
        if (PrimaryFile == null || !PrimaryFile.IsLoaded)
        {
            MessageBox.Show("请先加载主表文件", "参数错误", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        if (SecondaryFile == null || !SecondaryFile.IsLoaded)
        {
            MessageBox.Show("请先加载辅助表文件", "参数错误", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        if (PrimaryFile.SelectedWorksheets.Count == 0)
        {
            MessageBox.Show("请选择至少一个主表工作表", "参数错误", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        if (SecondaryFile.SelectedWorksheets.Count == 0)
        {
            MessageBox.Show("请选择至少一个辅助表工作表", "参数错误", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        if (SelectedPrimaryMatchFields.Count == 0 || SelectedSecondaryMatchFields.Count == 0)
        {
            MessageBox.Show("请选择匹配字段", "参数错误", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        if (SelectedPrimaryMatchFields.Count != SelectedSecondaryMatchFields.Count)
        {
            MessageBox.Show("主表和辅助表的匹配字段数量必须相等", "参数错误", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        return true;
    }

    /// <summary>
    ///     验证和更新字段映射
    /// </summary>
    private void ValidateAndUpdateFieldMappings()
    {
        // 处理主表匹配字段
        var validPrimaryMatchFields = new ObservableCollection<string>();
        foreach (var field in SelectedPrimaryMatchFields)
            if (PrimaryColumns.Contains(field))
                validPrimaryMatchFields.Add(field);

        SelectedPrimaryMatchFields.Clear();
        foreach (var field in validPrimaryMatchFields) SelectedPrimaryMatchFields.Add(field);

        // 处理辅助表匹配字段
        var validSecondaryMatchFields = new ObservableCollection<string>();
        foreach (var field in SelectedSecondaryMatchFields)
            if (SecondaryColumns.Contains(field))
                validSecondaryMatchFields.Add(field);

        SelectedSecondaryMatchFields.Clear();
        foreach (var field in validSecondaryMatchFields) SelectedSecondaryMatchFields.Add(field);

        // 处理字段映射
        var validMappings = new List<FieldMapping>();
        foreach (var mapping in FieldMappings)
            // 检查源字段是否仍然存在
            if (!string.IsNullOrEmpty(mapping.SourceField) && SecondaryColumns.Contains(mapping.SourceField))
                validMappings.Add(mapping);

        FieldMappings.Clear();
        foreach (var mapping in validMappings)
        {
            // 必须重新添加PropertyChanged事件处理程序
            ((INotifyPropertyChanged)mapping).PropertyChanged += FieldMapping_PropertyChanged;
            FieldMappings.Add(mapping);
        }
    }

    // 移除字段映射
    private void RemoveFieldMapping(FieldMapping? mapping)
    {
        if (mapping != null)
        {
            // 移除属性变更监听
            ((INotifyPropertyChanged)mapping).PropertyChanged -= FieldMapping_PropertyChanged;

            FieldMappings.Remove(mapping);
            NotifyCommandsCanExecuteChanged();
        }
    }

    // 添加主表筛选条件
    private void AddPrimaryFilter()
    {
        if (PrimaryColumns.Count == 0) return;

        try
        {
            // 创建新的筛选条件
            var filter = new FilterCondition
            {
                Field = PrimaryColumns.FirstOrDefault() ?? string.Empty,
                Operator = FilterOperator.Equals,
                Value = string.Empty,
                LogicalOperator = LogicalOperator.And
            };

            // 监听属性变更以自动刷新预览
            ((INotifyPropertyChanged)filter).PropertyChanged += FilterCondition_PropertyChanged;

            // 添加到筛选列表
            PrimaryFilters.Add(filter);
            Debug.WriteLine($"添加主表筛选条件: {filter.Field} {filter.Operator} {filter.Value}");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"添加主表筛选条件时出错: {ex.Message}");
            MessageBox.Show($"添加主表筛选条件时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // 判断是否可以添加主表筛选条件
    private bool CanAddPrimaryFilter()
    {
        var canAdd = PrimaryColumns != null && PrimaryColumns.Count > 0 && !IsBusy;
        Debug.WriteLine($"CanAddPrimaryFilter: {canAdd}, PrimaryColumns数量: {PrimaryColumns?.Count ?? 0}");
        return canAdd;
    }

    // 移除主表筛选条件
    private void RemovePrimaryFilter(FilterCondition? filter)
    {
        if (filter != null)
        {
            // 移除属性变更监听
            ((INotifyPropertyChanged)filter).PropertyChanged -= FilterCondition_PropertyChanged;

            PrimaryFilters.Remove(filter);
        }
    }

    // 添加辅助表筛选条件
    private void AddSecondaryFilter()
    {
        if (SecondaryColumns.Count == 0) return;

        try
        {
            // 创建新的筛选条件
            var filter = new FilterCondition
            {
                Field = SecondaryColumns.FirstOrDefault() ?? string.Empty,
                Operator = FilterOperator.Equals,
                Value = string.Empty,
                LogicalOperator = LogicalOperator.And
            };

            // 监听属性变更以自动刷新预览
            ((INotifyPropertyChanged)filter).PropertyChanged += FilterCondition_PropertyChanged;

            // 添加到筛选列表
            SecondaryFilters.Add(filter);
            Debug.WriteLine($"添加辅助表筛选条件: {filter.Field} {filter.Operator} {filter.Value}");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"添加辅助表筛选条件时出错: {ex.Message}");
            MessageBox.Show($"添加辅助表筛选条件时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // 判断是否可以添加辅助表筛选条件
    private bool CanAddSecondaryFilter()
    {
        var canAdd = SecondaryColumns != null && SecondaryColumns.Count > 0 && !IsBusy;
        Debug.WriteLine($"CanAddSecondaryFilter: {canAdd}, SecondaryColumns数量: {SecondaryColumns?.Count ?? 0}");
        return canAdd;
    }

    // 移除辅助表筛选条件
    private void RemoveSecondaryFilter(FilterCondition? filter)
    {
        if (filter != null)
        {
            // 移除属性变更监听
            ((INotifyPropertyChanged)filter).PropertyChanged -= FilterCondition_PropertyChanged;

            SecondaryFilters.Remove(filter);
        }
    }

    /// <summary>
    ///     开始数据合并操作
    /// </summary>
    private async Task StartMergeAsync()
    {
        if (!ValidateMergeParameters())
            return;

        try
        {
            var startTime = DateTime.Now;
            IsBusy = true;
            StatusMessage = "准备开始合并数据...";
            ProgressValue = 0;

            // 强制刷新数据确保使用最新内容
            var needRefresh = false;

            // 检查文件是否有更新
            if (File.Exists(PrimaryFilePath))
            {
                var currentLastWrite = File.GetLastWriteTime(PrimaryFilePath);
                if (PrimaryFile == null || currentLastWrite > PrimaryFile.LastChecked)
                {
                    needRefresh = true;
                    LogDebug($"检测到主表文件变更: {currentLastWrite} > {PrimaryFile?.LastChecked}");
                }
            }

            if (File.Exists(SecondaryFilePath))
            {
                var currentLastWrite = File.GetLastWriteTime(SecondaryFilePath);
                if (SecondaryFile == null || currentLastWrite > SecondaryFile.LastChecked)
                {
                    needRefresh = true;
                    LogDebug($"检测到辅助表文件变更: {currentLastWrite} > {SecondaryFile?.LastChecked}");
                }
            }

            if (needRefresh)
            {
                // 显示文件变更提示
                var refreshContent = new StackPanel { Margin = new Thickness(24) };
                refreshContent.Children.Add(new TextBlock
                {
                    Text = "检测到文件变更",
                    FontSize = 18,
                    FontWeight = FontWeights.Medium,
                    Margin = new Thickness(0, 0, 0, 16)
                });

                refreshContent.Children.Add(new TextBlock
                {
                    Text = "Excel文件已在外部被修改。是否刷新数据后再继续合并？",
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 0, 0, 24)
                });

                var buttonPanel1 = new StackPanel
                    { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
                var continueButton = new Button
                {
                    Content = "直接继续",
                    Style = Application.Current.Resources["MaterialDesignOutlinedButton"] as Style,
                    Margin = new Thickness(0, 0, 8, 0)
                };

                var refreshButton = new Button
                {
                    Content = "刷新后合并",
                    Style = Application.Current.Resources["MaterialDesignOutlinedLightButton"] as Style,
                    IsDefault = true
                };

                buttonPanel1.Children.Add(continueButton);
                buttonPanel1.Children.Add(refreshButton);
                refreshContent.Children.Add(buttonPanel1);

                var refreshFirst = false;
                continueButton.Click += (s, e) => { DialogHost.Close("RootDialog"); };
                refreshButton.Click += (s, e) =>
                {
                    refreshFirst = true;
                    DialogHost.Close("RootDialog");
                };

                await DialogHost.Show(refreshContent, "RootDialog");

                if (refreshFirst)
                {
                    // 先刷新数据
                    await RefreshDataAsync();

                    // 再次验证合并参数
                    if (!ValidateMergeParameters())
                        return;
                }
            }

            // 设置进度报告回调
            var progress = new Progress<(int Current, int Total, string Message)>(report =>
            {
                ProgressValue = report.Current;
                ProgressMaximum = report.Total;
                StatusMessage = report.Message;
            });

            // 根据选择的工作表数量选择不同的合并方法
            var result = PrimaryFile.SelectedWorksheets.Count > 1 || SecondaryFile.SelectedWorksheets.Count > 1
                ? await _excelFileManager.MergeMultipleWorksheetsAsync(
                    PrimaryFile,
                    SecondaryFile,
                    SelectedPrimaryMatchFields.ToList(),
                    SelectedSecondaryMatchFields.ToList(),
                    FieldMappings.ToList(),
                    PrimaryFilters.ToList(),
                    SecondaryFilters.ToList(),
                    progress)
                : await _excelFileManager.MergeExcelFilesAsync(
                    PrimaryFile,
                    SecondaryFile,
                    SelectedPrimaryMatchFields.ToList(),
                    SelectedSecondaryMatchFields.ToList(),
                    FieldMappings.ToList(),
                    PrimaryFilters.ToList(),
                    SecondaryFilters.ToList(),
                    progress);

            var endTime = DateTime.Now;
            var duration = endTime - startTime;

            // 显示结果对话框
            var successContent = new StackPanel { Margin = new Thickness(24) };
            successContent.Children.Add(new TextBlock
            {
                Text = "合并完成！",
                FontSize = 18,
                FontWeight = FontWeights.Medium,
                Margin = new Thickness(0, 0, 0, 16)
            });


            var resultGrid = new Grid();
            resultGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Auto) });
            resultGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            resultGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            resultGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            resultGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            resultGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            resultGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });


            // 添加结果信息
            AddResultRow(resultGrid, 0, "处理记录数:", result.ProcessedRows.ToString());
            AddResultRow(resultGrid, 1, "匹配记录数:", result.MatchedRows.ToString());
            AddResultRow(resultGrid, 2, "新增列数:", result.NewColumnsAdded.ToString());
            AddResultRow(resultGrid, 3, "结果已保存到:", result.OutputPath);
            AddResultRow(resultGrid, 4, "处理耗时:", FormatDuration(duration));

            successContent.Children.Add(resultGrid);

            // 添加按钮
            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 24, 0, 0)
            };
            var openFileButton = new Button
            {
                Content = "打开文件位置",
                Style = Application.Current.Resources["MaterialDesignOutlinedButton"] as Style,
                Margin = new Thickness(0, 0, 8, 0)
            };
            var closeButton = new Button
            {
                Content = "关闭",
                Style = Application.Current.Resources["MaterialDesignOutlinedLightButton"] as Style,
                IsDefault = true
            };

            openFileButton.Click += (s, e) =>
            {
                try
                {
                    // 打开文件所在文件夹并选中文件
                    var argument = "/select,\"" + result.OutputPath + "\"";
                    Process.Start("explorer.exe", argument);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"打开文件位置时出错: {ex.Message}");
                }
                finally
                {
                    DialogHost.Close("RootDialog");
                }
            };

            closeButton.Click += (s, e) => { DialogHost.Close("RootDialog"); };

            buttonPanel.Children.Add(openFileButton);
            buttonPanel.Children.Add(closeButton);
            successContent.Children.Add(buttonPanel);

            // 显示成功对话框
            await DialogHost.Show(successContent, "RootDialog");

            StatusMessage = "合并完成";

            // 更新文件的LastChecked属性以记录最新检查时间
            if (PrimaryFile != null && File.Exists(PrimaryFilePath))
                PrimaryFile.LastChecked = File.GetLastWriteTime(PrimaryFilePath);

            if (SecondaryFile != null && File.Exists(SecondaryFilePath))
                SecondaryFile.LastChecked = File.GetLastWriteTime(SecondaryFilePath);
        }
        catch (Exception ex)
        {
            // 提取内部异常信息以便用户更容易理解
            var errorMessage = ex.Message;
            var innerEx = ex.InnerException;
            while (innerEx != null)
            {
                errorMessage += $"\n详细信息: {innerEx.Message}";
                innerEx = innerEx.InnerException;
            }

            // 解析常见错误并提供更友好的提示
            if (errorMessage.Contains("does not belong to table"))
            {
                // 提取列名
                var columnName = string.Empty;
                var startIndex = errorMessage.IndexOf("Column '") + 8;
                if (startIndex > 8)
                {
                    var endIndex = errorMessage.IndexOf("'", startIndex);
                    if (endIndex > startIndex) columnName = errorMessage.Substring(startIndex, endIndex - startIndex);
                }

                if (!string.IsNullOrEmpty(columnName))
                    errorMessage = $"合并失败: 在处理字段 '{columnName}' 时出错，此字段在某些工作表中不存在。\n\n" +
                                   "请检查您的字段映射，确保所有映射的源字段都存在于选定的辅助表工作表中。";
                else
                    errorMessage = "合并失败: 尝试访问不存在的列。请检查您的字段映射，确保所有映射的源字段都存在于选定的辅助表工作表中。";
            }
            else if (errorMessage.Contains("being used by another process"))
            {
                errorMessage = "合并失败: 文件被其他程序占用。请关闭可能打开了这些Excel文件的程序后再试。";
            }

            // 显示错误对话框
            var errorContent = new StackPanel { Margin = new Thickness(24) };
            errorContent.Children.Add(new TextBlock
            {
                Text = "合并数据时出错",
                FontSize = 18,
                FontWeight = FontWeights.Medium,
                Foreground = new SolidColorBrush(Colors.Red),
                Margin = new Thickness(0, 0, 0, 16)
            });

            errorContent.Children.Add(new TextBlock
            {
                Text = errorMessage,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 24)
            });

            var errorButtonPanel = new StackPanel
                { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var retryButton = new Button
            {
                Content = "重试",
                Style = Application.Current.Resources["MaterialDesignOutlinedButton"] as Style,
                Margin = new Thickness(0, 0, 8, 0),
                Command = new RelayCommand(() =>
                {
                    DialogHost.Close("RootDialog");
                    // 延迟执行，确保对话框已关闭
                    Application.Current.Dispatcher.BeginInvoke(new Action(async () =>
                    {
                        await Task.Delay(100);
                        await StartMergeAsync();
                    }));
                })
            };

            var okButton = new Button
            {
                Content = "确定",
                Style = Application.Current.Resources["MaterialDesignOutlinedLightButton"] as Style,
                IsDefault = true
            };

            okButton.Click += (s, e) => { DialogHost.Close("RootDialog"); };
            errorButtonPanel.Children.Add(retryButton);
            errorButtonPanel.Children.Add(okButton);
            errorContent.Children.Add(errorButtonPanel);

            await DialogHost.Show(errorContent, "RootDialog");

            StatusMessage = "合并数据失败";
            Debug.WriteLine($"合并数据异常: {ex}");
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>
    ///     格式化时间
    /// </summary>
    /// <param name="duration"></param>
    /// <returns></returns>
    private string FormatDuration(TimeSpan duration)
    {
        if (duration.TotalHours >= 1)
            return $"{duration.Hours}小时{duration.Minutes}分钟{duration.Seconds}秒";
        if (duration.TotalMinutes >= 1)
            return $"{duration.Minutes}分钟{duration.Seconds}秒";
        if (duration.TotalSeconds >= 1)
            return $"{duration.Seconds}.{duration.Milliseconds:000}秒";
        return $"{duration.Milliseconds}毫秒";
    }

    /// <summary>
    ///     记录调试信息
    /// </summary>
    private void LogDebug(string message)
    {
        Debug.WriteLine($"{DateTime.Now:HH:mm:ss.fff} - {message}");
    }

    /// <summary>
    ///     向结果网格中添加一行信息
    /// </summary>
    private void AddResultRow(Grid grid, int rowIndex, string label, string value)
    {
        var labelBlock = new TextBlock
        {
            Text = label,
            FontWeight = FontWeights.SemiBold,
            Margin = new Thickness(0, 4, 16, 4),
            VerticalAlignment = VerticalAlignment.Center
        };

        var valueBlock = new TextBlock
        {
            Text = value,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 4, 0, 4),
            VerticalAlignment = VerticalAlignment.Center
        };

        grid.Children.Add(labelBlock);
        Grid.SetRow(labelBlock, rowIndex);
        Grid.SetColumn(labelBlock, 0);

        grid.Children.Add(valueBlock);
        Grid.SetRow(valueBlock, rowIndex);
        Grid.SetColumn(valueBlock, 1);
    }

    /// <summary>
    ///     检查文件是否有变更
    /// </summary>
    private Task<bool> CheckFilesForChangesAsync()
    {
        return Task.Run(() =>
        {
            try
            {
                var hasChanges = false;

                // 检查主表文件
                if (!string.IsNullOrEmpty(PrimaryFilePath) && File.Exists(PrimaryFilePath))
                {
                    var lastWriteTime = File.GetLastWriteTime(PrimaryFilePath);
                    if (PrimaryFile != null && PrimaryFile.LastChecked < lastWriteTime) hasChanges = true;
                }

                // 检查辅助表文件
                if (!string.IsNullOrEmpty(SecondaryFilePath) && File.Exists(SecondaryFilePath))
                {
                    var lastWriteTime = File.GetLastWriteTime(SecondaryFilePath);
                    if (SecondaryFile != null && SecondaryFile.LastChecked < lastWriteTime) hasChanges = true;
                }

                return hasChanges;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"检查文件变更时出错: {ex.Message}");
                return false;
            }
        });
    }

    // 验证合并参数
    private bool ValidateMergeParameters()
    {
        // 验证文件是否已加载
        if (PrimaryFile == null || !PrimaryFile.IsLoaded)
        {
            MessageBox.Show("请先加载主表文件", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        if (SecondaryFile == null || !SecondaryFile.IsLoaded)
        {
            MessageBox.Show("请先加载辅助表文件", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        // 验证是否选择了工作表
        if (SelectedPrimaryWorksheets.Count == 0)
        {
            MessageBox.Show("请选择至少一个主表工作表", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        if (SelectedSecondaryWorksheets.Count == 0)
        {
            MessageBox.Show("请选择至少一个辅助表工作表", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        // 验证匹配字段
        if (SelectedPrimaryMatchFields.Count == 0)
        {
            MessageBox.Show("请至少选择一个主表匹配字段", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        if (SelectedSecondaryMatchFields.Count == 0)
        {
            MessageBox.Show("请至少选择一个辅助表匹配字段", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        if (SelectedPrimaryMatchFields.Count != SelectedSecondaryMatchFields.Count)
        {
            MessageBox.Show("主表和辅助表的匹配字段数量必须相同", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        // 验证字段映射
        if (FieldMappings.Count == 0)
        {
            MessageBox.Show("请至少添加一个字段映射", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        foreach (var mapping in FieldMappings)
        {
            if (string.IsNullOrEmpty(mapping.SourceField))
            {
                MessageBox.Show("字段映射中的源字段不能为空", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            if (string.IsNullOrEmpty(mapping.TargetField))
            {
                MessageBox.Show("字段映射中的目标字段不能为空", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
        }

        return true;
    }

    // 判断是否可以开始合并
    private bool CanStartMerge()
    {
        return PrimaryFile != null && PrimaryFile.IsLoaded &&
               SecondaryFile != null && SecondaryFile.IsLoaded &&
               SelectedPrimaryWorksheets.Count > 0 &&
               SelectedSecondaryWorksheets.Count > 0 &&
               SelectedPrimaryMatchFields.Count > 0 &&
               SelectedSecondaryMatchFields.Count > 0 &&
               FieldMappings.Count > 0 &&
               !IsBusy;
    }

    // 保存配置
    /// <summary>
    ///     保存配置
    /// </summary>
    private async Task SaveConfigurationAsync()
    {
        try
        {
            // 准备配置对象
            var config = CreateConfigurationFromCurrentState();

            // 创建Material Design 3风格的对话框内容
            var configNameContent = new StackPanel { Margin = new Thickness(24) };

            // 添加标题
            configNameContent.Children.Add(new TextBlock
            {
                Text = "保存配置",
                FontSize = 18,
                FontWeight = FontWeights.Medium,
                Margin = new Thickness(0, 0, 0, 24)
            });

            // 添加说明文字
            configNameContent.Children.Add(new TextBlock
            {
                Text = "请输入配置名称:",
                Margin = new Thickness(0, 0, 0, 16)
            });

            // 添加输入框
            var textBox = new TextBox
            {
                Text = config.Name,
                Style = Application.Current.Resources["MaterialDesignOutlinedTextBox"] as Style,
                Margin = new Thickness(0, 0, 0, 24)
            };
            HintAssist.SetHint(textBox, "配置名称");
            configNameContent.Children.Add(textBox);

            // 添加按钮
            var buttonPanel2 = new StackPanel
                { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var cancelButton = new Button
            {
                Content = "取消",
                Style = Application.Current.Resources["MaterialDesignOutlinedButton"] as Style,
                Margin = new Thickness(0, 0, 8, 0)
            };
            var saveButton = new Button
            {
                Content = "保存",
                Style = Application.Current.Resources["MaterialDesignRaisedButton"] as Style,
                IsDefault = true
            };
            buttonPanel2.Children.Add(cancelButton);
            buttonPanel2.Children.Add(saveButton);
            configNameContent.Children.Add(buttonPanel2);

            // 显示对话框
            var dialogResult = false;

            cancelButton.Click += (s, e) => { DialogHost.Close("RootDialog"); };
            saveButton.Click += (s, e) =>
            {
                dialogResult = true;
                DialogHost.Close("RootDialog");
            };

            await DialogHost.Show(configNameContent, "RootDialog");

            if (!dialogResult)
                return;

            config.Name = textBox.Text;
            if (string.IsNullOrEmpty(config.Name))
                config.Name = $"配置_{DateTime.Now:yyyyMMddHHmmss}";

            StatusMessage = "正在保存配置...";
            IsBusy = true;

            var filePath = await _configurationManager.SaveConfigurationAsync(config);

            // 显示成功消息
            var successContent = new StackPanel { Margin = new Thickness(24) };

            successContent.Children.Add(new TextBlock
            {
                Text = "保存成功",
                FontSize = 18,
                FontWeight = FontWeights.Medium,
                Margin = new Thickness(0, 0, 0, 16)
            });

            successContent.Children.Add(new TextBlock
            {
                Text = $"配置已保存到：{filePath}",
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 24)
            });

            var okButton = new Button
            {
                Content = "确定",
                Style = Application.Current.Resources["MaterialDesignOutlinedLightButton"] as Style,
                HorizontalAlignment = HorizontalAlignment.Right,
                IsDefault = true
            };

            okButton.Click += (s, e) => { DialogHost.Close("RootDialog"); };
            successContent.Children.Add(okButton);

            await DialogHost.Show(successContent, "RootDialog");

            StatusMessage = "配置已保存";
        }
        catch (Exception ex)
        {
            // 显示错误消息
            var errorContent = new StackPanel { Margin = new Thickness(24) };

            errorContent.Children.Add(new TextBlock
            {
                Text = "保存失败",
                FontSize = 18,
                FontWeight = FontWeights.Medium,
                Foreground = new SolidColorBrush(Colors.Red),
                Margin = new Thickness(0, 0, 0, 16)
            });

            errorContent.Children.Add(new TextBlock
            {
                Text = $"保存配置时出错: {ex.Message}",
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 24)
            });

            var okButton = new Button
            {
                Content = "确定",
                Style = Application.Current.Resources["MaterialDesignOutlinedLightButton"] as Style,
                HorizontalAlignment = HorizontalAlignment.Right,
                IsDefault = true
            };

            okButton.Click += (s, e) => { DialogHost.Close("RootDialog"); };
            errorContent.Children.Add(okButton);

            await DialogHost.Show(errorContent, "RootDialog");

            StatusMessage = "保存配置失败";
            Debug.WriteLine($"保存配置异常: {ex}");
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>
    ///     加载配置
    /// </summary>
    /// <summary>
    ///     加载配置
    /// </summary>
    private async Task LoadConfigurationAsync()
    {
        try
        {
            StatusMessage = "正在获取配置列表...";
            IsBusy = true;

            var configurations = await _configurationManager.GetAllConfigurationsAsync();

            if (configurations.Count == 0)
            {
                // 显示无配置消息
                var noConfigContent = new StackPanel { Margin = new Thickness(24) };

                noConfigContent.Children.Add(new TextBlock
                {
                    Text = "没有可用配置",
                    FontSize = 18,
                    FontWeight = FontWeights.Medium,
                    Margin = new Thickness(0, 0, 0, 16)
                });

                noConfigContent.Children.Add(new TextBlock
                {
                    Text = "没有找到任何保存的配置。请先保存一个配置。",
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 0, 0, 24)
                });

                var buttonPanel1 = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Right
                };

                var openDirButton1 = new Button
                {
                    Content = "打开配置目录",
                    Style = Application.Current.Resources["MaterialDesignOutlinedButton"] as Style,
                    Margin = new Thickness(0, 0, 8, 0)
                };

                var okButton = new Button
                {
                    Content = "确定",
                    Style = Application.Current.Resources["MaterialDesignOutlinedLightButton"] as Style,
                    IsDefault = true
                };

                openDirButton1.Click += (s, e) =>
                {
                    OpenConfigurationDirectory();
                    DialogHost.Close("RootDialog");
                };

                okButton.Click += (s, e) => { DialogHost.Close("RootDialog"); };

                buttonPanel1.Children.Add(openDirButton1);
                buttonPanel1.Children.Add(okButton);
                noConfigContent.Children.Add(buttonPanel1);

                await DialogHost.Show(noConfigContent, "RootDialog");
                return;
            }

            // 创建配置选择对话框
            var dialogContent = new StackPanel { Margin = new Thickness(24) };

            // 添加标题
            dialogContent.Children.Add(new TextBlock
            {
                Text = "选择配置",
                FontSize = 18,
                FontWeight = FontWeights.Medium,
                Margin = new Thickness(0, 0, 0, 24)
            });

            // 添加配置列表容器
            var cardContainer = new Card
            {
                Style = Application.Current.Resources["MD3ListContainer"] as Style,
                Margin = new Thickness(0, 0, 0, 24),
                MaxHeight = 300
            };

            // 添加配置列表
            var listBox = new ListBox
            {
                Style = Application.Current.Resources["MD3ListBox"] as Style,
                SelectionMode = SelectionMode.Single
            };

            foreach (var configuration in configurations)
            {
                var configItem = new ListBoxItem
                {
                    Tag = configuration.Path
                };

                // 创建配置项内容
                var itemGrid = new Grid();
                itemGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                itemGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                var contentPanel = new StackPanel { Orientation = Orientation.Vertical };

                var titleBlock = new TextBlock
                {
                    Text = configuration.Name,
                    FontWeight = FontWeights.Medium,
                    Margin = new Thickness(0, 0, 0, 4)
                };

                var dateBlock = new TextBlock
                {
                    Text = $"创建于 {File.GetCreationTime(configuration.Path):yyyy-MM-dd HH:mm:ss}",
                    FontSize = 12,
                    Foreground = new SolidColorBrush(Colors.Gray)
                };

                contentPanel.Children.Add(titleBlock);
                contentPanel.Children.Add(dateBlock);

                // 删除按钮
                var deleteButton = new Button
                {
                    Style = Application.Current.Resources["MaterialDesignIconButton"] as Style,
                    ToolTip = "删除此配置",
                    Margin = new Thickness(8, 0, 0, 0)
                };

                var deleteIcon = new PackIcon
                {
                    Kind = PackIconKind.DeleteOutline,
                    Width = 20,
                    Height = 20,
                    Foreground = new SolidColorBrush(Colors.Red)
                };

                deleteButton.Content = deleteIcon;

                // 删除按钮事件处理
                deleteButton.Click += async (s, e) =>
                {
                    e.Handled = true; // 防止触发ListBoxItem选择

                    // 执行删除操作（使用NestedDialog避免冲突）
                    await DeleteConfigurationAsync(configuration);

                    // 如果删除成功，重新加载配置列表
                    if (!File.Exists(configuration.Path)) // 文件已被删除
                    {
                        // 关闭当前对话框
                        DialogHost.Close("RootDialog");

                        // 等待对话框完全关闭后重新打开
                        await Task.Delay(100);
                        await LoadConfigurationAsync();
                    }
                };

                Grid.SetColumn(contentPanel, 0);
                Grid.SetColumn(deleteButton, 1);

                itemGrid.Children.Add(contentPanel);
                itemGrid.Children.Add(deleteButton);

                configItem.Content = itemGrid;
                listBox.Items.Add(configItem);
            }

            // 选择第一项
            if (listBox.Items.Count > 0)
                listBox.SelectedIndex = 0;

            cardContainer.Content = listBox;
            dialogContent.Children.Add(cardContainer);

            // 添加按钮
            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            var openDirButton = new Button
            {
                Content = "打开目录",
                Style = Application.Current.Resources["MaterialDesignOutlinedButton"] as Style,
                Margin = new Thickness(0, 0, 8, 0)
            };

            var cancelButton = new Button
            {
                Content = "取消",
                Style = Application.Current.Resources["MaterialDesignOutlinedButton"] as Style,
                Margin = new Thickness(0, 0, 8, 0)
            };

            var selectButton = new Button
            {
                Content = "选择",
                Style = Application.Current.Resources["MaterialDesignOutlinedLightButton"] as Style,
                IsDefault = true
            };

            buttonPanel.Children.Add(openDirButton);
            buttonPanel.Children.Add(cancelButton);
            buttonPanel.Children.Add(selectButton);
            dialogContent.Children.Add(buttonPanel);

            // 显示对话框
            var dialogResult = false;

            openDirButton.Click += (s, e) => { OpenConfigurationDirectory(); };

            cancelButton.Click += (s, e) => { DialogHost.Close("RootDialog"); };
            selectButton.Click += (s, e) =>
            {
                dialogResult = true;
                DialogHost.Close("RootDialog");
            };

            await DialogHost.Show(dialogContent, "RootDialog");

            if (!dialogResult || listBox.SelectedItem == null)
                return;

            var configPath = ((ListBoxItem)listBox.SelectedItem).Tag as string;
            if (string.IsNullOrEmpty(configPath))
                return;

            StatusMessage = "正在加载配置...";
            var config = await _configurationManager.LoadConfigurationAsync(configPath);

            // 应用配置
            await ApplyConfigurationAsync(config);

            StatusMessage = $"配置 '{config.Name}' 已加载";
        }
        catch (Exception ex)
        {
            // 显示错误消息
            var errorContent = new StackPanel { Margin = new Thickness(24) };

            errorContent.Children.Add(new TextBlock
            {
                Text = "加载失败",
                FontSize = 18,
                FontWeight = FontWeights.Medium,
                Foreground = new SolidColorBrush(Colors.Red),
                Margin = new Thickness(0, 0, 0, 16)
            });

            errorContent.Children.Add(new TextBlock
            {
                Text = $"加载配置时出错: {ex.Message}",
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 24)
            });

            var okButton = new Button
            {
                Content = "确定",
                Style = Application.Current.Resources["MaterialDesignOutlinedLightButton"] as Style,
                HorizontalAlignment = HorizontalAlignment.Right,
                IsDefault = true
            };

            okButton.Click += (s, e) => { DialogHost.Close("RootDialog"); };
            errorContent.Children.Add(okButton);

            await DialogHost.Show(errorContent, "RootDialog");

            StatusMessage = "加载配置失败";
            Debug.WriteLine($"加载配置异常: {ex}");
        }
        finally
        {
            IsBusy = false;
        }
    }

    // 重置配置
    private void ResetConfiguration()
    {
        // 使用MD3风格对话框确认重置
        var dialogContent = new StackPanel { Margin = new Thickness(16) };
        dialogContent.Children.Add(new TextBlock
        {
            Text = "确定要重置所有配置吗？",
            FontSize = 16,
            Margin = new Thickness(0, 0, 0, 8)
        });
        dialogContent.Children.Add(new TextBlock
        {
            Text = "这将清除当前的所有设置。",
            Opacity = 0.7,
            Margin = new Thickness(0, 0, 0, 16)
        });

        var buttonPanel = new StackPanel
            { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        var cancelButton = new Button
        {
            Content = "取消", Style = Application.Current.Resources["MaterialDesignOutlinedButton"] as Style,
            Margin = new Thickness(8, 0, 0, 0)
        };
        var confirmButton = new Button
            { Content = "确定", Style = Application.Current.Resources["MaterialDesignRaisedButton"] as Style };
        buttonPanel.Children.Add(cancelButton);
        buttonPanel.Children.Add(confirmButton);
        dialogContent.Children.Add(buttonPanel);

        // 显示对话框
        var dialogResult = false;

        // 使用Material Design的DialogHost
        cancelButton.Click += (s, e) => { DialogHost.Close("RootDialog"); };
        confirmButton.Click += (s, e) =>
        {
            dialogResult = true;
            DialogHost.Close("RootDialog");
        };

        DialogHost.Show(dialogContent, "RootDialog").ContinueWith(t =>
        {
            if (dialogResult)
                Application.Current.Dispatcher.Invoke(() =>
                {
                    try
                    {
                        // 清空文件路径和密码
                        PrimaryFilePath = string.Empty;
                        PrimaryFilePassword = string.Empty;
                        SecondaryFilePath = string.Empty;
                        SecondaryFilePassword = string.Empty;

                        // 重置文件对象
                        PrimaryFile = new ExcelFile();
                        SecondaryFile = new ExcelFile();

                        // 清空集合
                        PrimaryWorksheets.Clear();
                        SecondaryWorksheets.Clear();
                        PrimaryColumns.Clear();
                        SecondaryColumns.Clear();
                        SelectedPrimaryMatchFields.Clear();
                        SelectedSecondaryMatchFields.Clear();
                        SelectedPrimaryWorksheets.Clear();
                        SelectedSecondaryWorksheets.Clear();
                        FieldMappings.Clear();
                        PrimaryFilters.Clear();
                        SecondaryFilters.Clear();

                        // 清空预览数据
                        PrimaryPreviewData = null;
                        SecondaryPreviewData = null;

                        // 重置状态
                        StatusMessage = "已重置所有配置";
                        ProgressValue = 0;

                        // 通知命令状态更新
                        NotifyCommandsCanExecuteChanged();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"重置配置时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                        Debug.WriteLine($"重置配置异常: {ex}");
                    }
                });
        });
    }

    /// <summary>
    ///     从当前状态创建配置对象
    /// </summary>
    private Configuration CreateConfigurationFromCurrentState()
    {
        var config = new Configuration
        {
            Name = $"配置_{DateTime.Now:yyyyMMddHHmmss}",
            PrimaryFilePath = PrimaryFilePath,
            PrimaryFilePassword = PrimaryFilePassword,
            SecondaryFilePath = SecondaryFilePath,
            SecondaryFilePassword = SecondaryFilePassword
        };

        // 处理工作表信息
        if (PrimaryFile != null)
        {
            config.PrimaryWorksheet = PrimaryFile.SelectedWorksheet;
            config.PrimaryWorksheets = new List<string>(SelectedPrimaryWorksheets);
        }

        if (SecondaryFile != null)
        {
            config.SecondaryWorksheet = SecondaryFile.SelectedWorksheet;
            config.SecondaryWorksheets = new List<string>(SelectedSecondaryWorksheets);
        }

        // 处理字段匹配和映射
        config.PrimaryMatchFields = new List<string>(SelectedPrimaryMatchFields);
        config.SecondaryMatchFields = new List<string>(SelectedSecondaryMatchFields);

        // 深度复制字段映射
        config.FieldMappings = FieldMappings.Select(m => new FieldMapping
        {
            SourceField = m.SourceField,
            TargetField = m.TargetField
        }).ToList();

        // 深度复制筛选条件
        config.PrimaryFilters = PrimaryFilters.Select(f => new FilterCondition
        {
            Field = f.Field,
            Operator = f.Operator,
            Value = f.Value,
            LogicalOperator = f.LogicalOperator
        }).ToList();

        config.SecondaryFilters = SecondaryFilters.Select(f => new FilterCondition
        {
            Field = f.Field,
            Operator = f.Operator,
            Value = f.Value,
            LogicalOperator = f.LogicalOperator
        }).ToList();

        return config;
    }

    /// <summary>
    ///     应用配置
    /// </summary>
    private async Task ApplyConfigurationAsync(Configuration config)
    {
        if (config == null)
            return;

        try
        {
            // 重置现有配置前先记录当前状态
            IsBusy = true;
            StatusMessage = $"正在应用配置: {config.Name}";

            // 设置文件路径和密码
            PrimaryFilePath = config.PrimaryFilePath;
            PrimaryFilePassword = config.PrimaryFilePassword;
            SecondaryFilePath = config.SecondaryFilePath;
            SecondaryFilePassword = config.SecondaryFilePassword;

            // 加载文件
            if (!string.IsNullOrEmpty(PrimaryFilePath) && File.Exists(PrimaryFilePath)) await LoadPrimaryFileAsync();

            if (!string.IsNullOrEmpty(SecondaryFilePath) && File.Exists(SecondaryFilePath))
                await LoadSecondaryFileAsync();

            // 设置工作表 - 必须等文件加载完成
            if (PrimaryFile != null && PrimaryFile.IsLoaded)
            {
                SelectedPrimaryWorksheets.Clear();

                // 先检查多表配置
                if (config.PrimaryWorksheets != null && config.PrimaryWorksheets.Count > 0)
                {
                    foreach (var worksheet in config.PrimaryWorksheets)
                        if (PrimaryWorksheets.Contains(worksheet))
                            SelectedPrimaryWorksheets.Add(worksheet);
                }
                // 再检查单表配置(向后兼容)
                else if (!string.IsNullOrEmpty(config.PrimaryWorksheet) &&
                         PrimaryWorksheets.Contains(config.PrimaryWorksheet))
                {
                    SelectedPrimaryWorksheets.Add(config.PrimaryWorksheet);
                }

                // 如果有选择的工作表，加载工作表信息
                if (SelectedPrimaryWorksheets.Count > 0) await LoadPrimaryWorksheetsInfoAsync();
            }

            if (SecondaryFile != null && SecondaryFile.IsLoaded)
            {
                SelectedSecondaryWorksheets.Clear();

                // 先检查多表配置
                if (config.SecondaryWorksheets != null && config.SecondaryWorksheets.Count > 0)
                {
                    foreach (var worksheet in config.SecondaryWorksheets)
                        if (SecondaryWorksheets.Contains(worksheet))
                            SelectedSecondaryWorksheets.Add(worksheet);
                }
                // 再检查单表配置(向后兼容)
                else if (!string.IsNullOrEmpty(config.SecondaryWorksheet) &&
                         SecondaryWorksheets.Contains(config.SecondaryWorksheet))
                {
                    SelectedSecondaryWorksheets.Add(config.SecondaryWorksheet);
                }

                // 如果有选择的工作表，加载工作表信息
                if (SelectedSecondaryWorksheets.Count > 0) await LoadSecondaryWorksheetsInfoAsync();
            }

            // 设置匹配字段 - 必须等工作表信息加载完成
            if (config.PrimaryMatchFields != null && PrimaryColumns.Count > 0)
            {
                SelectedPrimaryMatchFields.Clear();
                foreach (var field in config.PrimaryMatchFields)
                    if (PrimaryColumns.Contains(field))
                        SelectedPrimaryMatchFields.Add(field);
            }

            if (config.SecondaryMatchFields != null && SecondaryColumns.Count > 0)
            {
                SelectedSecondaryMatchFields.Clear();
                foreach (var field in config.SecondaryMatchFields)
                    if (SecondaryColumns.Contains(field))
                        SelectedSecondaryMatchFields.Add(field);
            }

            // 设置字段映射 - 必须等工作表信息加载完成
            if (config.FieldMappings != null)
            {
                FieldMappings.Clear();
                foreach (var mapping in config.FieldMappings)
                    // 源字段必须存在于辅助表中
                    if (!string.IsNullOrEmpty(mapping.SourceField) && SecondaryColumns.Contains(mapping.SourceField))
                    {
                        var newMapping = new FieldMapping
                        {
                            SourceField = mapping.SourceField,
                            TargetField = mapping.TargetField
                        };

                        // 监听属性变更
                        ((INotifyPropertyChanged)newMapping).PropertyChanged += FieldMapping_PropertyChanged;

                        // 添加到映射集合
                        FieldMappings.Add(newMapping);
                    }
            }

            // 设置筛选条件 - 必须等工作表信息加载完成
            if (config.PrimaryFilters != null)
            {
                PrimaryFilters.Clear();
                foreach (var filter in config.PrimaryFilters)
                    if (!string.IsNullOrEmpty(filter.Field) && PrimaryColumns.Contains(filter.Field))
                    {
                        var newFilter = new FilterCondition
                        {
                            Field = filter.Field,
                            Operator = filter.Operator,
                            Value = filter.Value,
                            LogicalOperator = filter.LogicalOperator
                        };

                        // 监听属性变更
                        ((INotifyPropertyChanged)newFilter).PropertyChanged += FilterCondition_PropertyChanged;

                        // 添加到筛选集合
                        PrimaryFilters.Add(newFilter);
                    }
            }

            if (config.SecondaryFilters != null)
            {
                SecondaryFilters.Clear();
                foreach (var filter in config.SecondaryFilters)
                    if (!string.IsNullOrEmpty(filter.Field) && SecondaryColumns.Contains(filter.Field))
                    {
                        var newFilter = new FilterCondition
                        {
                            Field = filter.Field,
                            Operator = filter.Operator,
                            Value = filter.Value,
                            LogicalOperator = filter.LogicalOperator
                        };

                        // 监听属性变更
                        ((INotifyPropertyChanged)newFilter).PropertyChanged += FilterCondition_PropertyChanged;

                        // 添加到筛选集合
                        SecondaryFilters.Add(newFilter);
                    }
            }

            // 加载预览数据
            await LoadPreviewDataWithFiltersAsync();

            // 通知命令状态更新
            NotifyCommandsCanExecuteChanged();
            StatusMessage = $"配置 '{config.Name}' 已应用";
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"应用配置时出错: {ex.Message}");
            StatusMessage = "应用配置失败";
            throw;
        }
        finally
        {
            IsBusy = false;
        }
    }

    #endregion
}