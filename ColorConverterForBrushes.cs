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
                case "Grey":
                    return Brushes.Gray;
                case "Blond":
                    return Brushes.Gold;
                case "Brown":
                    return Brushes.Brown;
                case "Dark Blue":
                    return Brushes.DarkBlue;
                case "Cyan":
                    return Brushes.Cyan;
                case "Ice":
                    return Brushes.AliceBlue;
                case "Red":
                    return Brushes.Red;
                case "Green":
                    return Brushes.Green;
                case "LimeGreen":
                    return Brushes.LimeGreen;
                case "Purple":
                    return Brushes.Purple;
                case "Pink":
                    return Brushes.Pink;
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
