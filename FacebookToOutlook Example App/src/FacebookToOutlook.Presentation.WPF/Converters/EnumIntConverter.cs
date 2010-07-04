using System;
using System.Windows.Data;

namespace FacebookToOutlook.Presentation.Converters
{
    public class EnumIntConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return (int)value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return Enum.GetValues(targetType).GetValue((int)value);
        }
    }
}
