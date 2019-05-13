using System.Windows;
using System.Windows.Media;

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
