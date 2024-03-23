using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;

namespace Szakdolgozat.Helper
{
    /// <summary>
    /// A helper class that provides methods to navigate the visual tree of WPF controls.
    /// </summary>
    internal class ChildParentHelper
    {
        /// <summary>
        /// Retrieves the first child of a specified type from a parent visual.
        /// If the child is not found directly under the parent, it recursively searches the visual tree.
        /// </summary>
        /// <typeparam name="T">The type of the child to find.</typeparam>
        /// <param name="parent">The parent visual.</param>
        /// <returns>The first child of the specified type, or null if no such child is found.</returns>
        private T GetVisualChild<T>(Visual parent) where T : Visual
        {
            T child = default(T);
            int numVisuals = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < numVisuals; i++)
            {
                Visual v = (Visual)VisualTreeHelper.GetChild(parent, i);
                child = v as T;
                if (child == null)
                {
                    child = GetVisualChild<T>(v);
                }
                if (child != null)
                {
                    break;
                }
            }
            return child;
        }

        /// <summary>
        /// Retrieves the first parent of a specified type from a child visual.
        /// If the parent is not found directly above the child, it recursively searches the visual tree.
        /// </summary>
        /// <typeparam name="T">The type of the parent to find.</typeparam>
        /// <param name="child">The child visual.</param>
        /// <returns>The first parent of the specified type, or null if no such parent is found.</returns>
        public T FindVisualParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parentObject = VisualTreeHelper.GetParent(child);

            if (parentObject == null)
                return null;

            T parent = parentObject as T;
            if (parent != null)
            {
                return parent;
            }
            else
            {
                return FindVisualParent<T>(parentObject);
            }
        }

        /// <summary>
        /// Retrieves a cell from a Datagrid at a specified column index for a given row.
        /// If the row is null or the DataGridCellPresenter cannot be found, it returns null.
        /// </summary>
        /// <param name="grid">The DataGrid from which to retrieve the cell.</param>
        /// <param name="row">The DataGridRow for which to retrieve the cell.</param>
        /// <param name="columnIndex">The index of the column in which the cell is located.</param>
        /// <returns>The DataGridCell at the specified column index for the given row, or null if it cannot be found.</returns>
        public DataGridCell GetCell(DataGrid grid, DataGridRow row, int columnIndex)
        {
            if (row != null)
            {
                DataGridCellsPresenter presenter = GetVisualChild<DataGridCellsPresenter>(row);
                if (presenter != null)
                {
                    return presenter.ItemContainerGenerator.ContainerFromIndex(columnIndex) as DataGridCell;
                }
            }

            return null;
        }
    }
}
