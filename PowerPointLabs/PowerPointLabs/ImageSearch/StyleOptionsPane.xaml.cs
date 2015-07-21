using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using PowerPointLabs.ImageSearch.Model;

namespace PowerPointLabs.ImageSearch
{
    /// <summary>
    /// Interaction logic for StyleOptionsPane.xaml
    /// </summary>
    public partial class StyleOptionsPane
    {
        public StyleOptionsPane()
        {
            InitializeComponent();
        }

        private void ColorPanel_OnMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            var panel = sender as Border;
            if (panel == null) return;

            var colorDialog = new ColorDialog
            {
                Color = GetColor(panel.Background as SolidColorBrush),
                FullOpen = true
            };
            if (colorDialog.ShowDialog() != DialogResult.OK) return;
            
            var hexString = GetHexValue(colorDialog.Color);
            var options = DataContext as StyleOptions;
            if (options != null)
            {
                switch (panel.Name)
                {
                    case "FontColorPanel":
                        options.FontColor = hexString;
                        break;
                    case "OverlayColorPanel":
                        options.OverlayColor = hexString;
                        break;
                }
            }
        }

        public System.Drawing.Color GetColor(SolidColorBrush brush)
        {
            return System.Drawing.Color.FromArgb(brush.Color.A, brush.Color.R, brush.Color.G, brush.Color.B);
        }

        public Color GetColor(System.Drawing.Color color)
        {
            return Color.FromArgb(color.A, color.R, color.G, color.B);
        }

        private string GetHexValue(System.Drawing.Color color)
        {
            byte[] rgbArray = { color.R, color.G, color.B };
            var hex = BitConverter.ToString(rgbArray);
            return "#" + hex.Replace("-", "");
        }

        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            var options = DataContext as StyleOptions;
            if (options != null)
            {
                options.Init();
            }
        }
    }
}
