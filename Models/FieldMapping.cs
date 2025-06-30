using System.ComponentModel;
using System.Runtime.CompilerServices;

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
            if (SetProperty(ref _sourceField, value))
                // 当源字段变更时，总是将目标字段更新为源字段的值
                if (!string.IsNullOrEmpty(value))
                    TargetField = value;
        }
    }

    /// <summary>
    ///     目标字段名称（主表中的字段）
    /// </summary>
    public string TargetField
    {
        get => _targetField;
        set => SetProperty(ref _targetField, value);
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
            return false;

        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }

    protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}