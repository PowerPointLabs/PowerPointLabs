using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using PowerPointLabs.Views;
using Drawing = System.Drawing;

namespace PowerPointLabs.EffectsLab.Views
{
    /// <summary>
    /// Interaction logic for SpotlightSettingsDialogBox.xaml
    /// </summary>
    public partial class SpotlightSettingsDialogBox
    {
        public delegate void DialogConfirmedDelegate(float spotlightTransparency, float spotlightSoftEdges, Drawing.Color spotlightColor);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }
        
        private float lastTransparency;

        public SpotlightSettingsDialogBox()
        {
            InitializeComponent();
        }

        public SpotlightSettingsDialogBox(float spotlightTransparency, float spotlightSoftEdges, Drawing.Color spotlightColor)
            : this()
        {
            lastTransparency = spotlightTransparency;
            spotlightTransparencyInput.Text = spotlightTransparency.ToString("P0");
            spotlightTransparencyInput.ToolTip = EffectsLabText.SettingsTransparencyInputTooltip;

            String[] keys = EffectsLabSettings.SpotlightSoftEdgesMapping.Keys.ToArray();
            float[] values = EffectsLabSettings.SpotlightSoftEdgesMapping.Values.ToArray();
            softEdgesSelectionInput.ItemsSource = keys;
            softEdgesSelectionInput.Content = keys[Array.IndexOf(values, spotlightSoftEdges)];
            softEdgesSelectionInput.ToolTip = EffectsLabText.SettingsSoftEdgesSelectionInputTooltip;

            spotlightColorRect.Fill = new SolidColorBrush(GraphicsUtil.MediaColorFromDrawingColor(spotlightColor));
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            ValidateSpotlightTransparencyInput();
            string text = spotlightTransparencyInput.Text;
            if (text.Contains("%"))
            {
                text = text.Substring(0, text.IndexOf("%"));
            }
            DialogConfirmedHandler(float.Parse(text) / 100,
                            EffectsLabSettings.SpotlightSoftEdgesMapping[(String)softEdgesSelectionInput.Content], 
                            GraphicsUtil.DrawingColorFromBrush(spotlightColorRect.Fill));
            Close();
        }

        private void SpotlightColorRect_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Color currentColor = (spotlightColorRect.Fill as SolidColorBrush).Color;
            WPFColorDialog colorDialog = new WPFColorDialog();
            colorDialog.Color = GraphicsUtil.DrawingColorFromMediaColor(currentColor);
            colorDialog.FullOpen = true;
            if (colorDialog.ShowDialog() != Utils.DialogResult.Cancel)
            {
                spotlightColorRect.Fill = GraphicsUtil.MediaBrushFromDrawingColor(colorDialog.Color);
            }
        }

        private void SoftEdgesSelectionInput_Item_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left && softEdgesSelectionInput.IsExpanded)
            {
                string value = ((TextBlock)e.Source).Text;
                softEdgesSelectionInput.Content = value;
            }
        }

        private void SpotlightTransparencyInput_LostFocus(object sender, RoutedEventArgs e)
        {
            ValidateSpotlightTransparencyInput();
        }

        private void ValidateSpotlightTransparencyInput()
        {
            float enteredValue;
            string text = spotlightTransparencyInput.Text;
            if (text.Contains("%"))
            {
                text = text.Substring(0, text.IndexOf("%"));
            }

            if (float.TryParse(text, out enteredValue) &&
                enteredValue > 0 && 
                enteredValue <= 100)
            {
                lastTransparency = enteredValue / 100;
            }
            spotlightTransparencyInput.Text = lastTransparency.ToString("P0");
        }
    }
}
