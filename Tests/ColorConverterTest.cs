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
        public void ColorTest()
        {
            string stringColorName = "Black";
            string expectedResult = "#000000";
            object colorName = (object)stringColorName;
            string result = _converter.ColorNameToHex(colorName);

            ClassicAssert.AreEqual(expectedResult, result);
        }
    }
}
