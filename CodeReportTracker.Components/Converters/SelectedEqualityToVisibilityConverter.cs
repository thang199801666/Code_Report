using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace CodeReportTracker.Components.Converters
{
    /// <summary>
    /// MultiValue converter: values[0] = item, values[1] = selectedItem.
    /// Returns Visible when Equals(item, selectedItem), otherwise Collapsed.
    /// </summary>
    public class SelectedEqualityToVisibilityConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values == null || values.Length < 2)
                return Visibility.Collapsed;

            var item = values[0];
            var selected = values[1];

            return Equals(item, selected) ? Visibility.Visible : Visibility.Collapsed;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }
}