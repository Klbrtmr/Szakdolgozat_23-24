using System;
using System.Globalization;
using System.Windows.Data;

namespace Szakdolgozat.Converters
{
    /// <summary>
    /// This class is used to convert a height value to a percentage of its original value.
    /// </summary>
    public class PercentageConverter : IValueConverter
    {
        /// <summary>
        /// Converts a height value to a percentage of its original value.
        /// </summary>
        /// <param name="value">The original height value as an object.</param>
        /// <param name="targetType">The type of the binding target property. This parameter will be ignored.</param>
        /// <param name="parameter">The percentage as a string.</param>
        /// <param name="culture">The culture to use in the converter. This parameter will be ignored.</param>
        /// <returns>The height value multiplied by the percentage, or the original value if the conversion is unsuccessful.</returns>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is double actualHeight && 
                parameter is string percentage &&
                double.TryParse(percentage, NumberStyles.Number, CultureInfo.InvariantCulture, out double percent))
            {
                return actualHeight * percent;
            }

            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
