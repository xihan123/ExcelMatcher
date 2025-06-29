using System.ComponentModel;

namespace ExcelMatcher.Models;

/// <summary>
///     字段映射模型，表示源字段与目标字段的映射关系
/// </summary>
public class FieldMapping : INotifyPropertyChanged
{
    private string _sourceField = string.Empty;
    private string _targetField = string.Empty;

    /// <summary>
    ///     源字段名称（辅助表中的字段）
    /// </summary>
    public string SourceField
    {
        get => _sourceField;
        set
        {
            _sourceField = value;
            OnPropertyChanged(nameof(SourceField));
        }
    }

    /// <summary>
    ///     目标字段名称（主表中的字段）
    /// </summary>
    public string TargetField
    {
        get => _targetField;
        set
        {
            _targetField = value;
            OnPropertyChanged(nameof(TargetField));
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    protected virtual void OnPropertyChanged(string propertyName)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}