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

namespace PowerPointLabs.ResizeLab
{
    /// <summary>
    /// Interaction logic for ResizePane.xaml
    /// </summary>
    public partial class ResizePaneWPF : UserControl
    {
        public ResizePaneWPF()
        {
            InitializeComponent();
        }

        private void SameWidthBtn_Click(object sender, RoutedEventArgs e)
        {
            ResizeLabMain.ResizeToSameWidth();
        }

        private void SameHeightBtn_Click(object sender, RoutedEventArgs e)
        {
            ResizeLabMain.ResizeToSameHeight();
        }

        private void SameSizeBtn_Click(object sender, RoutedEventArgs e)
        {
            ResizeLabMain.ResizeToSameHeightAndWidth();
        }
    }
}
