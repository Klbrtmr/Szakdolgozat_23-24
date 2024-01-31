using System;
using System.Globalization;
using System.Windows.Data;

namespace Szakdolgozat
{
    public class PercentageConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is double actualHeight && parameter is string percentage)
            {
                if (double.TryParse(percentage, NumberStyles.Number, CultureInfo.InvariantCulture, out double percent))
                {
                    return actualHeight * percent;
                }
            }

            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
