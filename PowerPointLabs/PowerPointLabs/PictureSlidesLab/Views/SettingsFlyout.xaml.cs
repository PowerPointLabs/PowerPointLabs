using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.Utils;

using Color = System.Drawing.Color;

namespace PowerPointLabs.PictureSlidesLab.Views
{
    /// <summary>
    /// Interaction logic for SettingsFlyout.xaml
    /// </summary>
    public partial class SettingsFlyout
    {
        public SettingsFlyout()
        {
            InitializeComponent();
            UpdateInsertCitationControlsVisibility();
        }

        public void UpdateInsertCitationControlsVisibility()
        {
            if (InsertCitationToggleSwitch.IsChecked != null
                && InsertCitationToggleSwitch.IsChecked.Value)
            {
                CitationFontColorLabel.Visibility = Visibility.Visible;
                CitationFontColorPanel.Visibility = Visibility.Visible;

                CitationFontSizeLabel.Visibility = Visibility.Visible;
                CitationFontSizeSlider.Visibility = Visibility.Visible;

                CitationAlignmentLabel.Visibility = Visibility.Visible;
                CitationAlignmentComboBox.Visibility = Visibility.Visible;

                UseTextBoxLabel.Visibility = Visibility.Visible;
                UseTextBoxToggleSwitch.Visibility = Visibility.Visible;

                CitationTextBoxColorLabel.Visibility = Visibility.Visible;
                CitationTextBoxColorPanel.Visibility = Visibility.Visible;
            }
            else
            {
                CitationFontColorLabel.Visibility = Visibility.Collapsed;
                CitationFontColorPanel.Visibility = Visibility.Collapsed;

                CitationFontSizeLabel.Visibility = Visibility.Collapsed;
                CitationFontSizeSlider.Visibility = Visibility.Collapsed;

                CitationAlignmentLabel.Visibility = Visibility.Collapsed;
                CitationAlignmentComboBox.Visibility = Visibility.Collapsed;

                UseTextBoxLabel.Visibility = Visibility.Collapsed;
                UseTextBoxToggleSwitch.Visibility = Visibility.Collapsed;

                CitationTextBoxColorLabel.Visibility = Visibility.Collapsed;
                CitationTextBoxColorPanel.Visibility = Visibility.Collapsed;
            }
        }

        private void ColorPanel_OnMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            Border panel = sender as Border;
            if (panel == null)
            {
                return;
            }

            ColorDialog colorDialog = new ColorDialog
            {
                Color = GetColor(panel.Background as SolidColorBrush),
                FullOpen = true
            };
            if (colorDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            string hexString = StringUtil.GetHexValue(colorDialog.Color);
            Settings settings = DataContext as Settings;
            if (settings == null)
            {
                return;
            }

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

        private void InsertCitationToggleSwitch_OnIsCheckedChanged(object sender, EventArgs e)
        {
            UpdateInsertCitationControlsVisibility();
        }
    }
}
