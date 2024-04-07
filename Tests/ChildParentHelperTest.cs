using NUnit.Framework;
using System.Windows.Controls;
using Szakdolgozat.Helper;
using NUnit.Framework.Legacy;

namespace Szakdolgozat.Tests
{
    [TestFixture]
    [Apartment(System.Threading.ApartmentState.STA)]
    internal class ChildParentHelperTest
    {
        private ChildParentHelper _helper;

        [SetUp]
        public void SetUp()
        {
            _helper = new ChildParentHelper();
        }

        [Test]
        public void FindVisualParent_ValidChild_ReturnsParent()
        {
            // Arrange
            var parent = new Grid();
            var child = new Button();
            parent.Children.Add(child);

            // Act
            var foundParent = _helper.FindVisualParent<Grid>(child);

            // Assert
            ClassicAssert.AreEqual(parent, foundParent);
        }

        [Test]
        public void FindVisualParent_InvalidChild_ReturnsNull()
        {
            // Arrange
            var child = new Button();

            // Act
            var foundParent = _helper.FindVisualParent<Grid>(child);

            // Assert
            ClassicAssert.IsNull(foundParent);
        }

        [Test]
        public void GetCell_NullRow_ReturnsNull()
        {
            // Arrange
            var grid = new DataGrid();

            // Act
            var cell = _helper.GetCell(grid, null, 0);

            // Assert
            ClassicAssert.IsNull(cell);
        }
    }
}
