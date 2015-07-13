using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace PowerPointLabs.ImageSearch
{
    /// <summary>
    /// Interaction logic for ImageSearchPane.xaml
    /// </summary>
    public partial class ImageSearchPane : UserControl
    {
        public ImageSearchPane()
        {
            InitializeComponent();
            SearchListBox.ItemsSource = new List<SearchResultItem>()
            {
                new SearchResultItem()
                {
                    ImageFile = @"C:\Users\Giki\Downloads\Marz-Z-at-Dreamer-Hack.jpg"
                }
            };
        }

        private void ListBoxItem_Selected(object sender, RoutedEventArgs e)
        {

        }

        private void SearchTextBox_OnMouseEnter(object sender, MouseEventArgs e)
        {
            SearchButton_OnClick(sender, e);
        }

        private void SearchButton_OnClick(object sender, RoutedEventArgs e)
        {
            
        }
    }
}
