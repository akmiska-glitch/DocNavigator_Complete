using Avalonia.Data.Converters;
using Avalonia.Media;
using System;
using System.Globalization;

namespace DocNavigator.App.Converters
{
    /// <summary>
    /// True -> Green, False -> Gray.
    /// </summary>
    public sealed class BoolToBrushConverter : IValueConverter
    {
        public IBrush TrueBrush { get; set; } = Brushes.LimeGreen;
        public IBrush FalseBrush { get; set; } = Brushes.DimGray;

        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value is bool b)
                return b ? TrueBrush : FalseBrush;
            return FalseBrush;
        }

        public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }
}
