using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;
using Szakdolgozat.Properties;

namespace Szakdolgozat.Converters
{
    internal class ColorConverterForBrushes : IValueConverter
    {
        public static readonly IValueConverter Instance = new ColorConverterForBrushes();

        /// <summary>
        /// A dictionary that maps color names to their corresponding Brush objects.
        /// </summary>
        private static readonly Dictionary<string, Brush> colorMap = new Dictionary<string, Brush>
        {
            {Resources.Black,      Brushes.Black },
            {Resources.White,      Brushes.White},
            {Resources.Gray,       Brushes.Gray},
            {Resources.Gold,       Brushes.Gold},
            {Resources.Brown,      Brushes.Brown},
            {Resources.Blue,       Brushes.Blue},
            {Resources.Cyan,       Brushes.Cyan},
            {Resources.AliceBlue,  Brushes.AliceBlue},
            {Resources.Red,        Brushes.Red},
            {Resources.Green,      Brushes.LawnGreen},
            {Resources.LimeGreen,  Brushes.LimeGreen},
            {Resources.Purple,     Brushes.Purple},
            {Resources.Pink,       Brushes.DeepPink},
            {Resources.Yellow,     Brushes.Yellow},
            {Resources.Orange,     Brushes.Orange}
        };


        /// <summary>
        /// Converts a color name to its corresponding Brush objects.
        /// If the color name is not found in the colorMap dictionary, it returns a Transparent brush.
        /// </summary>
        /// <param name="value">The color name as an object.</param>
        /// <param name="targetType">The type of the binding target property. This parameter will be ignored.</param>
        /// <param name="parameter">The converter parameter to use. This parameter will be ignored.</param>
        /// <param name="culture">The culture to use in the converter. This parameter will be ignored.</param>
        /// <returns>The corresponding Brush object for the color name.</returns>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string colorString &&
                colorMap.TryGetValue(colorString, out Brush brush))
            {
                return brush;
            }
            return Brushes.Transparent;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
