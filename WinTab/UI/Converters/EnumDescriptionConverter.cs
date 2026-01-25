using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Globalization;
using System.ComponentModel;

namespace WinTab.UI.Converters;

public class EnumDescriptionConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value == null) return DependencyProperty.UnsetValue;

        if (value is ComboBoxItem comboBoxItem)
            return comboBoxItem.Content ?? string.Empty;

        var valueStr = value.ToString()!;
        var fieldInfo = value.GetType().GetField(valueStr);
        if (fieldInfo == null) return valueStr;

        var descriptionAttribute = fieldInfo.GetCustomAttributes(typeof(DescriptionAttribute), false)
            .FirstOrDefault() as DescriptionAttribute;

        var mode = parameter?.ToString();
        if (string.Equals(mode, "Description", StringComparison.OrdinalIgnoreCase))
            return GetDescriptionValue(valueStr, descriptionAttribute);

        return GetDisplayValue(value, valueStr, descriptionAttribute);
    }

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        return value;
    }

    private static object GetDescriptionValue(string valueStr, DescriptionAttribute? descriptionAttribute)
    {
        var descriptionKey = descriptionAttribute?.Description;
        if (!string.IsNullOrWhiteSpace(descriptionKey) && Application.Current != null && Application.Current.Resources.Contains(descriptionKey))
            return Application.Current.FindResource(descriptionKey);

        return descriptionAttribute?.Description ?? valueStr;
    }

    private static object GetDisplayValue(object value, string valueStr, DescriptionAttribute? descriptionAttribute)
    {
        var resourceKey = $"HotKeyAction_{valueStr}";
        if (Application.Current != null && Application.Current.Resources.Contains(resourceKey))
            return Application.Current.FindResource(resourceKey);

        var typeKey = $"{value.GetType().Name}_{valueStr}";
        if (Application.Current != null && Application.Current.Resources.Contains(typeKey))
            return Application.Current.FindResource(typeKey);

        return descriptionAttribute?.Description ?? valueStr;
    }
}

