using System.ComponentModel;

namespace ExcelMatcher.Helpers;

/// <summary>
///     枚举辅助类，提供枚举值与显示文本的转换功能
/// </summary>
public static class EnumHelper
{
    /// <summary>
    ///     获取枚举值的描述文本
    /// </summary>
    /// <param name="value">枚举值</param>
    /// <returns>描述文本，如果未找到则返回枚举名称</returns>
    public static string GetDescription(Enum value)
    {
        var field = value.GetType().GetField(value.ToString());

        if (field != null &&
            Attribute.GetCustomAttribute(field, typeof(DescriptionAttribute)) is DescriptionAttribute attribute)
            return attribute.Description;

        return value.ToString();
    }

    /// <summary>
    ///     从描述文本获取枚举值
    /// </summary>
    /// <typeparam name="T">枚举类型</typeparam>
    /// <param name="description">描述文本</param>
    /// <returns>对应的枚举值，如果未找到则返回默认值</returns>
    public static T GetValueFromDescription<T>(string description) where T : Enum
    {
        foreach (var field in typeof(T).GetFields())
            if (Attribute.GetCustomAttribute(field, typeof(DescriptionAttribute)) is DescriptionAttribute attribute)
            {
                if (attribute.Description == description) return (T)field.GetValue(null);
            }
            else if (field.Name == description)
            {
                return (T)field.GetValue(null);
            }

        return default;
    }
}