using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace CodeReportTracker.Components.Converters
{
    // Simple multi-value converter used in the table XAML.
    public sealed class ValuesToColorConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values == null || values.Length < 3 || values[2] is not bool hasCheck || !hasCheck)
                return Brushes.White;

            var oldValue = values[0]?.ToString()?.Trim() ?? string.Empty;
            var newValue = values[1]?.ToString()?.Trim() ?? string.Empty;
            var hasUpdate = !string.Equals(oldValue, newValue, StringComparison.OrdinalIgnoreCase);

            return hasUpdate ? Brushes.Orange : Brushes.White;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
