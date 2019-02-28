using System;
using System.Collections.Generic;
using System.Drawing;
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

namespace PowerPointLabs.ColorsLab
{
    /// <summary>
    /// Interaction logic for ColorInformationDialog.xaml
    /// </summary>
    public partial class ColorInformationDialog : Window
    {

      
        public ColorInformationDialog(System.Drawing.Color color)
        {
            DataContext = this;

            WindowStartupLocation = WindowStartupLocation.CenterScreen;

            InitializeComponent();

            colorRectangle.Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(color.A, color.R, color.G, color.B));
            colorHexText.Text = ((HSLColor)color).ToHexString();
            colorRgbText.Text = ((HSLColor)color).ToRGBString();
            colorHslText.Text = ((HSLColor)color).ToString();

        }
    }
}
