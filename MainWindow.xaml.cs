using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Szakdolgozat
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void ListFiles()
        {
            string path = "../../Assets/";
            string[] files = Directory.GetFiles(path, "*.jpg");
            
            foreach (string file in files) 
            {
                filesListing.Items.Add(System.IO.Path.GetFileName(file));
            }

            if (filesListing.Items.Count == 0)
            {
                tabControlBorder.BorderBrush = Brushes.Red;
            }
            else
            {
                tabControlBorder.BorderBrush = Brushes.Green;
            }
        }

        private void filesListing_Loaded(object sender, RoutedEventArgs e)
        {
            ListFiles();

            List<UIElement> modifiedItems = new List<UIElement>();

            Random random = new Random();

            foreach (var item in filesListing.Items)
            {
                Ellipse ellipse = new Ellipse();
                ellipse.Width = 15;
                ellipse.Height = 15;

                Color randomColor = Color.FromRgb((byte)random.Next(256), (byte)random.Next(256), (byte)random.Next(256));
                SolidColorBrush brush = new SolidColorBrush(randomColor);

                ellipse.Fill = brush;

                StackPanel stackPanel = new StackPanel();
                stackPanel.Orientation = Orientation.Horizontal;
                stackPanel.Children.Add(ellipse);
                stackPanel.Children.Add(new TextBlock() { Text = item.ToString() });
                
                modifiedItems.Add(stackPanel);
            }

            filesListing.Items.Clear();

            foreach (var item in modifiedItems)
            {
                filesListing.Items.Add(item);
            }
        }
    }
}
