using System.Windows;
using ExcelMatcher.Services;
using ExcelMatcher.ViewModels;
using Microsoft.Extensions.DependencyInjection;

namespace ExcelMatcher;

/// <summary>
///     App.xaml 的交互逻辑
/// </summary>
public partial class App : Application
{
    private readonly ServiceProvider _serviceProvider;

    public App()
    {
        // 配置依赖注入
        var serviceCollection = new ServiceCollection();
        ConfigureServices(serviceCollection);
        _serviceProvider = serviceCollection.BuildServiceProvider();
    }

    private void ConfigureServices(IServiceCollection services)
    {
        // 注册服务
        services.AddSingleton<ExcelFileManager>();
        services.AddSingleton<ConfigurationManager>();

        // 注册视图模型
        services.AddSingleton<MainViewModel>();

        // 注册主窗口
        services.AddSingleton<MainWindow>();
    }

    private void Application_Startup(object sender, StartupEventArgs e)
    {
        // 获取主窗口实例
        var mainWindow = _serviceProvider.GetService<MainWindow>();

        // 设置数据上下文
        mainWindow.DataContext = _serviceProvider.GetService<MainViewModel>();

        // 显示主窗口
        mainWindow.Show();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        // 释放服务提供者
        _serviceProvider?.Dispose();

        base.OnExit(e);
    }
}