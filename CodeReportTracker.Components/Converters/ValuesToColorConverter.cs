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
            var s1 = values?.Length > 0 ? values[0] as string : null;
            var s2 = values?.Length > 1 ? values[1] as string : null;
            return string.Equals(s1, s2, StringComparison.Ordinal) ? Brushes.Transparent : Brushes.Orange;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}