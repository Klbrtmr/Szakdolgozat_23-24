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
            string[] files = Directory.GetFiles(path, "*.png");
            
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
        }
    }
}
