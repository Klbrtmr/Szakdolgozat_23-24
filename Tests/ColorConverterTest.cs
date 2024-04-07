using NUnit.Framework;
using NUnit.Framework.Internal;
using NUnit.Framework.Legacy;
using Szakdolgozat.Converters;

namespace Szakdolgozat.Tests
{
    [TestFixture]
    internal class ColorConverterTest
    {
        private ColorConverter _converter;

        [SetUp]
        public void SetUp()
        {
            _converter = new ColorConverter();
        }

        [Test]
        public void ColorTestForBlack()
        {
            // Arrange
            string stringColorName = "Black";
            string expectedResult = "#000000";
            object colorName = (object)stringColorName;

            // Act
            string result = _converter.ColorNameToHex(colorName);

            // Assert
            ClassicAssert.AreEqual(expectedResult, result);
        }

        [Test]
        public void ColorTestForPurple()
        {
            // Arrange
            string stringColorName = "Purple";
            string expectedResult = "#800080";
            object colorName = (object)stringColorName;

            // Act
            string result = _converter.ColorNameToHex(colorName);

            // Assert
            ClassicAssert.AreEqual(expectedResult, result);
        }
    }
}
