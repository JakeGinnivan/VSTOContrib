using System;
using System.Globalization;
using System.Windows.Data;

namespace FacebookToOutlook.Presentation.Converters
{
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms", MessageId = "Flags")]
    public class FlagsEnumConverter : IValueConverter
    {
        private int _targetValue;

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null) return false;
            var mask = (int)parameter;
            _targetValue = (int)value;
            return ((mask & _targetValue) != 0);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            _targetValue ^= (int)parameter;
            return Enum.Parse(targetType, _targetValue.ToString());
        }
    }
}
