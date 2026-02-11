using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace CodeReportTracker.Converters
{
    public class StatusToBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // Example: value is expected to be an enum or string representing status
            // Adjust logic as needed for your actual status values
            if (value == null)
                return Brushes.Gray;

            string status = value.ToString();
            switch (status)
            {
                case "Loaded":
                    return Brushes.Green;
                case "Loading":
                    return Brushes.Orange;
                case "Error":
                    return Brushes.Red;
                default:
                    return Brushes.Gray;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}