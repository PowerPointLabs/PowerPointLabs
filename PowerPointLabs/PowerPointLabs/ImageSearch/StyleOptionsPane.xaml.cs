using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using MahApps.Metro.Controls;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.Utils;
using Color = System.Drawing.Color;
using ComboBox = System.Windows.Controls.ComboBox;

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
            
            var hexString = StringUtil.GetHexValue(colorDialog.Color);
            var options = DataContext as StyleOptions;
            if (options == null) return;
            
            switch (panel.Name)
            {
                case "FontColorPanel":
                    options.FontColor = hexString;
                    break;
                case "OverlayColorPanel":
                    options.OverlayColor = hexString;
                    break;
                case "TextBoxOverlayColorPanel":
                    options.TextBoxOverlayColor = hexString;
                    break;
                case "BannerOverlayColorPanel":
                    options.BannerOverlayColor = hexString;
                    break;
            }
        }

        public Color GetColor(SolidColorBrush brush)
        {
            return Color.FromArgb(brush.Color.A, brush.Color.R, brush.Color.G, brush.Color.B);
        }

        private void ResetButton_OnClick(object sender, RoutedEventArgs e)
        {
            var options = DataContext as StyleOptions;
            if (options != null)
            {
                options.Init();
            }
        }

        private void ToggleSwitch_OnIsCheckedChanged(object sender, EventArgs e)
        {
            var toggleSwitch = sender as ToggleSwitch;
            if (toggleSwitch == null || toggleSwitch.IsChecked == null) return;

            if (toggleSwitch.IsChecked.Value)
            {
                FontFamilyComboBox.IsEnabled = false;
                FontSizeTextBox.IsEnabled = false;
                FontColorPanel.IsEnabled = false;
            }
            else
            {
                FontFamilyComboBox.IsEnabled = true;
                FontSizeTextBox.IsEnabled = true;
                FontColorPanel.IsEnabled = true;
            }
        }

        private void TextboxPositionComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var textboxPosCombobox = sender as ComboBox;
            if (textboxPosCombobox == null) return;

            var textboxPos = textboxPosCombobox.SelectedIndex;
            if (textboxPos == 0 /*Original*/)
            {
                TextboxAlignmentComboBox.IsEnabled = false;
            }
            else
            {
                TextboxAlignmentComboBox.IsEnabled = true;
            }
        }
    }
}
