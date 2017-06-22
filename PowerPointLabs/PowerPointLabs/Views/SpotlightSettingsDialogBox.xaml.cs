using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Media;

using PowerPointLabs.Utils;

namespace PowerPointLabs.Views
{
    /// <summary>
    /// Interaction logic for SpotlightSettingsDialogBox.xaml
    /// </summary>
    public partial class SpotlightSettingsDialogBox
    {
        public delegate void UpdateSettingsDelegate(float spotlightTransparency, float softEdge, System.Drawing.Color newColor);
        public UpdateSettingsDelegate SettingsHandler { get; set; }

        private Dictionary<String, float> softEdgesMapping = new Dictionary<string, float>
        {
            {"None", 0},
            {"1 Point", 1},
            {"2.5 Points", 2.5f},
            {"5 Points", 5},
            {"10 Points", 10},
            {"25 Points", 25},
            {"50 Points", 50}
        };
        private float lastTransparency;

        public SpotlightSettingsDialogBox()
        {
            InitializeComponent();
        }

        public SpotlightSettingsDialogBox(float defaultTransparency, float defaultSoftEdge, System.Drawing.Color defaultColor)
            : this()
        {
            lastTransparency = defaultTransparency;
            spotlightTransparencyInput.Text = defaultTransparency.ToString("P0");
            spotlightTransparencyInput.ToolTip = TextCollection.SpotlightSettingsTransparencyInputTooltip;

            String[] keys = softEdgesMapping.Keys.ToArray();
            float[] values = softEdgesMapping.Values.ToArray();
            softEdgesSelectionInput.ItemsSource = keys;
            softEdgesSelectionInput.SelectedIndex = Array.IndexOf(values, defaultSoftEdge);
            softEdgesSelectionInput.ToolTip = TextCollection.SpotlightSettingsSoftEdgesSelectionInputTooltip;

            spotlightColorRect.Fill = new SolidColorBrush(Graphics.MediaColorFromDrawingColor(defaultColor));
        }

        private void OkButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            SpotlightTransparencyInput_LostFocus(null, null);
            string text = spotlightTransparencyInput.Text;
            if (text.Contains("%"))
            {
                text = text.Substring(0, text.IndexOf("%"));
            }
            SettingsHandler(float.Parse(text) / 100, 
                            softEdgesMapping[(String)softEdgesSelectionInput.SelectedItem], 
                            Graphics.DrawingColorFromBrush(spotlightColorRect.Fill));
            Close();
        }

        private void CancelButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Close();
        }

        private void SpotlightTransparencyInput_LostFocus(object sender, System.Windows.RoutedEventArgs e)
        {
            float enteredValue;
            string text = spotlightTransparencyInput.Text;
            if (text.Contains("%"))
            {
                text = text.Substring(0, text.IndexOf("%"));
            }

            if (float.TryParse(text, out enteredValue))
            {
                if (enteredValue > 0 && enteredValue <= 100)
                {
                    lastTransparency = enteredValue / 100;
                }
                spotlightTransparencyInput.Text = lastTransparency.ToString("P0");
            }
            else
            {
                spotlightTransparencyInput.Text = lastTransparency.ToString("P0");
            }
        }

        private void SpotlightColorRect_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            Color currentColor = (spotlightColorRect.Fill as SolidColorBrush).Color;
            ColorDialog colorDialog = new ColorDialog();
            colorDialog.Color = Graphics.DrawingColorFromMediaColor(currentColor);
            colorDialog.FullOpen = true;
            if (colorDialog.ShowDialog() != System.Windows.Forms.DialogResult.Cancel)
            {
                spotlightColorRect.Fill = Graphics.MediaBrushFromDrawingColor(colorDialog.Color);
            }
        }
    }
}
