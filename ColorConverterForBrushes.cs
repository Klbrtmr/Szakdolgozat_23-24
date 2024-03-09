using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace Szakdolgozat
{
    internal class ColorConverterForBrushes : IValueConverter
    {
        public static readonly IValueConverter Instance = new ColorConverterForBrushes();

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string colorString = value.ToString();

            switch (colorString)
            {
                case "Black":
                    return Brushes.Black;
                case "White":
                    return Brushes.White;
                case "Gray":
                    return Brushes.Gray;
                case "Gold":
                    return Brushes.Gold;
                case "Brown":
                    return Brushes.Brown;
                case "Blue":
                    return Brushes.Blue;
                case "Cyan":
                    return Brushes.Cyan;
                case "Alice Blue":
                    return Brushes.AliceBlue;
                case "Red":
                    return Brushes.Red;
                case "Green":
                    return Brushes.LawnGreen;
                case "LimeGreen":
                    return Brushes.LimeGreen;
                case "Purple":
                    return Brushes.Purple;
                case "Pink":
                    return Brushes.DeepPink;
                case "Yellow":
                    return Brushes.Yellow;
                case "Orange":
                    return Brushes.Orange;
                default:
                    return Brushes.Transparent;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
