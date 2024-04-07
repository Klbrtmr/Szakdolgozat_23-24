using NUnit.Framework;
using NUnit.Framework.Legacy;
using System.Globalization;
using Szakdolgozat.Converters;

namespace Szakdolgozat.Tests
{
    [TestFixture]
    internal class PercentageConverterTest
    {
        private PercentageConverter _percentageConverter;

        [SetUp]
        public void SetUp()
        {
            _percentageConverter = new PercentageConverter();
        }

        [Test]
        public void Convert_ValidInput_ReturnsConvertedValue()
        {
            // Arrange
            double originalValue = 100.0;
            double expectedConvertedValue = 50.0;
            string percentage = "0.5"; // 50%

            // Act
            var convertedValue = _percentageConverter.Convert(originalValue, typeof(double), percentage, CultureInfo.InvariantCulture);

            // Assert
            ClassicAssert.AreEqual(expectedConvertedValue, convertedValue);
        }

        [Test]
        public void Convert_InvalidInput_ReturnsOriginalValue()
        {
            // Arrange
            double originalValue = 100.0;
            string invalidPercentage = "not_a_number";

            // Act
            var convertedValue = _percentageConverter.Convert(originalValue, typeof(double), invalidPercentage, CultureInfo.InvariantCulture);

            // Assert
            ClassicAssert.AreEqual(originalValue, convertedValue);
        }
    }
}
