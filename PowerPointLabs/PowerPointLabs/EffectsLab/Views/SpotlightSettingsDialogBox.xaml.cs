using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;

using PowerPointLabs.Utils;

using Drawing = System.Drawing;
using Forms = System.Windows.Forms;

namespace PowerPointLabs.EffectsLab.Views
{
    /// <summary>
    /// Interaction logic for SpotlightSettingsDialogBox.xaml
    /// </summary>
    public partial class SpotlightSettingsDialogBox
    {
        public delegate void DialogConfirmedDelegate(float spotlightTransparency, float softEdge, Drawing.Color newColor);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }

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

        public SpotlightSettingsDialogBox(float defaultTransparency, float defaultSoftEdge, Drawing.Color defaultColor)
            : this()
        {
            lastTransparency = defaultTransparency;
            spotlightTransparencyInput.Text = defaultTransparency.ToString("P0");
            spotlightTransparencyInput.ToolTip = TextCollection.SpotlightSettingsTransparencyInputTooltip;

            String[] keys = softEdgesMapping.Keys.ToArray();
            float[] values = softEdgesMapping.Values.ToArray();
            softEdgesSelectionInput.ItemsSource = keys;
            softEdgesSelectionInput.Content = keys[Array.IndexOf(values, defaultSoftEdge)];
            softEdgesSelectionInput.ToolTip = TextCollection.SpotlightSettingsSoftEdgesSelectionInputTooltip;

            spotlightColorRect.Fill = new SolidColorBrush(Graphics.MediaColorFromDrawingColor(defaultColor));
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
                            softEdgesMapping[(String)softEdgesSelectionInput.Content], 
                            Graphics.DrawingColorFromBrush(spotlightColorRect.Fill));
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void SpotlightColorRect_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Color currentColor = (spotlightColorRect.Fill as SolidColorBrush).Color;
            ColorDialog colorDialog = new ColorDialog();
            colorDialog.Color = Graphics.DrawingColorFromMediaColor(currentColor);
            colorDialog.FullOpen = true;
            if (colorDialog.ShowDialog() != Forms.DialogResult.Cancel)
            {
                spotlightColorRect.Fill = Graphics.MediaBrushFromDrawingColor(colorDialog.Color);
            }
        }

        void SoftEdgesSelectionInput_Item_MouseUp(object sender, MouseButtonEventArgs e)
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
