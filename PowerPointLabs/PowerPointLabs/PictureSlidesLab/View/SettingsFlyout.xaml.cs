using System;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.Utils;
using Color = System.Drawing.Color;

namespace PowerPointLabs.PictureSlidesLab.View
{
    /// <summary>
    /// Interaction logic for SettingsFlyout.xaml
    /// </summary>
    public partial class SettingsFlyout
    {
        public SettingsFlyout()
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

            var hexString = StringUtil.GetHexValue(colorDialog.Color);
            var settings = DataContext as Settings;
            if (settings == null) return;

            if (panel.Name == "CitationTextBoxColorPanel")
            {
                settings.CitationTextBoxColor = hexString;
            }
            else if (panel.Name == "CitationFontColorPanel")
            {
                settings.CitationFontColor = hexString;
            }
        }

        private Color GetColor(SolidColorBrush brush)
        {
            return Color.FromArgb(brush.Color.A, brush.Color.R, brush.Color.G, brush.Color.B);
        }
    }
}
