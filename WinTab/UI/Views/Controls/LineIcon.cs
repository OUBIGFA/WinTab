using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace WinTab.UI.Views.Controls;

/// <summary>A theme-aware Lucide stroke icon on its original 24-unit canvas.</summary>
public sealed class LineIcon : Control
{
    public static readonly DependencyProperty DataProperty = DependencyProperty.Register(
        nameof(Data), typeof(Geometry), typeof(LineIcon), new PropertyMetadata(null));

    public Geometry? Data
    {
        get => (Geometry?)GetValue(DataProperty);
        set => SetValue(DataProperty, value);
    }
}
