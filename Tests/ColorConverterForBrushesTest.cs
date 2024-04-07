using NUnit.Framework;
using NUnit.Framework.Legacy;
using System.Globalization;
using System.Windows.Media;
using Szakdolgozat.Converters;

namespace Szakdolgozat.Tests
{
    [TestFixture]
    internal class ColorConverterForBrushesTest
    {
        private ColorConverterForBrushes _converter;

        [SetUp]
        public void SetUp()
        {
            _converter = new ColorConverterForBrushes();
        }

        [Test]
        public void CanConvert()
        {
            // Arrange
            string stringColorName = "Cyan";
            var expectedResult = Brushes.Cyan;
            object colorName = (object)stringColorName;

            // Act
            var result = _converter.Convert(colorName, typeof(Brush), null, CultureInfo.CurrentCulture);

            // Assert
            ClassicAssert.AreEqual(expectedResult, result);
        }

    }
}
