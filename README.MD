# Excel数据匹配合并工具

![版本](https://img.shields.io/github/v/release/xihan123/ExcelMatcher)
![语言](https://img.shields.io/badge/语言-C%23-brightgreen.svg)
![许可证](https://img.shields.io/badge/许可证-MIT-orange.svg)
![下载量](https://img.shields.io/github/downloads/xihan123/ExcelMatcher/total)

一个功能强大的WPF应用程序，专门用于合并和匹配Excel文件中的数据。该工具支持多工作表处理、灵活的字段映射、数据筛选和配置管理等功能。

## 🚀 主要功能

### 📁 文件管理

- **多格式支持**：支持 `.xlsx` 和 `.xls` 格式的Excel文件
- **密码保护**：支持加载带密码保护的Excel文件
- **拖拽操作**：支持文件拖拽到指定区域快速加载
- **实时预览**：文件加载后实时显示工作表和数据预览

### 📊 工作表处理

- **多工作表选择**：支持同时选择多个工作表进行处理
- **智能合并**：自动合并来自不同工作表的数据
- **列信息展示**：自动识别并显示所有可用列

### 🔗 字段配置

- **灵活匹配**：支持多字段组合匹配（主表字段 ↔ 辅助表字段）
- **字段映射**：自定义辅助表字段到主表字段的映射关系
- **新增字段**：支持将辅助表数据作为新列添加到主表
- **智能提示**：提供字段选择的下拉菜单和自动完成

### 🔍 数据筛选

- **多条件筛选**：支持为主表和辅助表设置多个筛选条件
- **丰富操作符**：
    - 等于、不等于
    - 包含、不包含
    - 开始于、结束于
    - 大于、小于、大于等于、小于等于
    - 为空、不为空
- **逻辑运算**：支持 AND/OR 逻辑运算符组合条件
- **实时预览**：筛选条件应用后实时更新数据预览

### ⚙️ 配置管理

- **保存配置**：将当前所有设置保存为配置文件
- **加载配置**：快速加载之前保存的配置
- **配置管理**：查看、删除已保存的配置
- **一键重置**：快速清空所有配置重新开始

### 🔧 高级功能

- **数据诊断**：智能诊断匹配字段的问题和数据质量
- **进度跟踪**：详细的处理进度显示和状态反馈
- **错误处理**：友好的错误提示和异常处理
- **文件刷新**：检测文件变更并支持数据刷新

## 🎨 界面特色

- **Material Design 3**：采用最新的Material Design 3设计语言
- **现代化界面**：清新简洁的用户界面，操作直观
- **响应式布局**：支持窗口缩放和不同分辨率适配
- **深色/浅色主题**：根据系统设置自动适配主题
- **动画效果**：流畅的过渡动画和视觉反馈

## 🛠️ 技术栈

- **框架**：WPF (.NET 6+)
- **UI库**：Material Design In XAML Toolkit
- **MVVM框架**：CommunityToolkit.Mvvm
- **Excel处理**：专用Excel文件管理服务
- **依赖注入**：Microsoft.Extensions.DependencyInjection
- **数据序列化**：Newtonsoft.Json

## 📋 系统要求

- **操作系统**：Windows 10 1903 或更高版本
- **.NET版本**：.NET 6.0 或更高版本
- **内存**：建议 4GB 以上
- **存储空间**：50MB 可用空间

## 🚀 快速开始

### 安装运行

1. 下载最新版本的安装包
2. 运行安装程序完成安装
3. 启动 "Excel数据匹配合并工具"

### 基本使用流程

1. **加载文件**
    - 点击"浏览"按钮或直接拖拽Excel文件到对应区域
    - 输入文件密码（如果需要）
    - 等待文件加载完成

2. **选择工作表**
    - 从主表和辅助表的工作表列表中选择需要处理的工作表
    - 支持多选（按住Ctrl键）

3. **配置匹配字段**
    - 在左侧选择主表的匹配字段
    - 在右侧选择辅助表的匹配字段
    - 确保两侧字段数量相同且对应关系正确

4. **设置字段映射**
    - 点击"添加字段映射"
    - 选择辅助表的源字段
    - 指定或输入主表的目标字段名

5. **应用筛选条件**（可选）
    - 展开筛选条件面板
    - 添加所需的筛选条件
    - 预览筛选后的数据

6. **开始合并**
    - 检查配置无误后点击"开始合并"
    - 等待处理完成
    - 查看合并结果并打开输出文件

## 💡 使用技巧

### 最佳实践

1. **数据准备**
    - 确保匹配字段的数据格式一致
    - 建议使用唯一标识符作为匹配字段
    - 清理数据中的多余空格和特殊字符

2. **性能优化**
    - 大文件处理前建议先用小样本测试配置
    - 合理使用筛选条件减少处理数据量
    - 定期清理不需要的配置文件

3. **错误处理**
    - 遇到问题时使用"诊断"功能分析原因
    - 保存工作配置以便快速重试
    - 查看详细的错误信息和建议

### 常用快捷操作

- **Ctrl + 拖拽**：多选工作表或字段
- **双击标题栏**：最大化/还原窗口

## 🔧 配置文件

配置文件自动保存在：

```
%LOCALAPPDATA%\ExcelMatcher\Configurations\
```

配置文件包含：

- 文件路径和密码设置
- 工作表选择
- 字段匹配和映射关系
- 筛选条件
- 其他自定义设置

## 🐛 故障排除

### 常见问题

**Q: 文件加载失败**

- 检查文件是否被其他程序占用
- 确认文件密码是否正确
- 验证文件格式是否受支持

**Q: 匹配结果不正确**

- 使用"诊断"功能检查匹配字段
- 确认字段数据类型和格式
- 检查是否存在重复值

**Q: 处理速度慢**

- 使用筛选条件减少数据量
- 关闭其他占用内存的程序
- 考虑分批处理大文件

**Q: 界面显示异常**

- 检查系统显示缩放设置
- 更新显卡驱动程序
- 重启应用程序

## 🔄 更新日志

### 版本特性

- 支持多工作表处理
- Material Design 3界面设计
- 增强的错误处理和用户反馈
- 配置管理功能
- 数据诊断工具

## 🤝 贡献

欢迎提交Issue和Pull Request来改进这个项目！

### 开发环境搭建

1. 克隆仓库

```bash
git clone https://github.com/xihan123/ExcelMatcher.git
```

2. 打开解决方案

```bash
cd ExcelMatcher
# 使用 Visual Studio 2022 或 JetBrains Rider 打开 ExcelMatcher.sln
```

3. 还原NuGet包

```bash
dotnet restore
```

4. 构建项目

```bash
dotnet build
```

## 🤝 如何贡献

我们欢迎任何形式的贡献！

1. Fork 本仓库
2. 创建您的特性分支 (`git checkout -b feature/amazing-feature`)
3. 提交您的更改 (`git commit -m '添加一些很棒的功能'`)
4. 推送到分支 (`git push origin feature/amazing-feature`)
5. 打开一个 Pull Request

## 📜 许可证

该项目采用 MIT 许可证 - 详情请参见 [LICENSE](LICENSE) 文件

## 👨‍💻 作者

[xihan123](https://github.com/xihan123)

## 📊 项目状态

![GitHub stars](https://img.shields.io/github/stars/xihan123/ExcelMatcher?style=social)
![GitHub forks](https://img.shields.io/github/forks/xihan123/ExcelMatcher?style=social)
![GitHub issues](https://img.shields.io/github/issues/xihan123/ExcelMatcher)
![GitHub pull requests](https://img.shields.io/github/issues-pr/xihan123/ExcelMatcher)

**⭐ 如果这个工具对您有帮助，请给项目点个Star支持！**
