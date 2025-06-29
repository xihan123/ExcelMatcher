using System.ComponentModel;

namespace ExcelMatcher.Models;

/// <summary>
///     筛选操作符枚举
/// </summary>
public enum FilterOperator
{
    [Description("等于")] Equals,
    [Description("不等于")] NotEquals,
    [Description("包含")] Contains,
    [Description("不包含")] NotContains,
    [Description("开始于")] StartsWith,
    [Description("结束于")] EndsWith,
    [Description("大于")] GreaterThan,
    [Description("小于")] LessThan,
    [Description("大于等于")] GreaterThanOrEqual,
    [Description("小于等于")] LessThanOrEqual,
    [Description("为空")] IsNull,
    [Description("不为空")] IsNotNull
}

/// <summary>
///     逻辑运算符枚举
/// </summary>
public enum LogicalOperator
{
    [Description("与")] And,
    [Description("或")] Or
}

/// <summary>
///     筛选条件模型类，表示一条筛选规则
/// </summary>
public class FilterCondition : INotifyPropertyChanged
{
    private string _field = string.Empty;
    private LogicalOperator _logicalOperator = LogicalOperator.And;
    private FilterOperator _operator = FilterOperator.Equals;
    private string _value = string.Empty;

    /// <summary>
    ///     字段名称
    /// </summary>
    public string Field
    {
        get => _field;
        set
        {
            _field = value;
            OnPropertyChanged(nameof(Field));
        }
    }

    /// <summary>
    ///     筛选操作符
    /// </summary>
    public FilterOperator Operator
    {
        get => _operator;
        set
        {
            _operator = value;
            OnPropertyChanged(nameof(Operator));
        }
    }

    /// <summary>
    ///     筛选值
    /// </summary>
    public string Value
    {
        get => _value;
        set
        {
            _value = value;
            OnPropertyChanged(nameof(Value));
        }
    }

    /// <summary>
    ///     逻辑运算符（与前一个条件的关系）
    /// </summary>
    public LogicalOperator LogicalOperator
    {
        get => _logicalOperator;
        set
        {
            _logicalOperator = value;
            OnPropertyChanged(nameof(LogicalOperator));
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    protected virtual void OnPropertyChanged(string propertyName)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}