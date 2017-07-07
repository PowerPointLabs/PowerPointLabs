﻿using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;

using PowerPointLabs.Utils;

using Drawing = System.Drawing;
using Forms = System.Windows.Forms;

namespace PowerPointLabs.HighlightLab.Views
{
    /// <summary>
    /// Interaction logic for HighlightLabSettingsDialogBox.xaml
    /// </summary>
    public partial class HighlightLabSettingsDialogBox
    {
        public delegate void UpdateSettingsDelegate(Drawing.Color highlightColor, 
                                                    Drawing.Color defaultColor, 
                                                    Drawing.Color backgroundColor);
        public UpdateSettingsDelegate SettingsHandler { get; set; }

        public HighlightLabSettingsDialogBox()
        {
            InitializeComponent();
        }

        public HighlightLabSettingsDialogBox(Drawing.Color defaultHighlightColor, 
                                            Drawing.Color defaultTextColor, 
                                            Drawing.Color defaultBackgroundColor)
            : this()
        {
            textHighlightColorRect.Fill = new SolidColorBrush(Graphics.MediaColorFromDrawingColor(defaultHighlightColor));
            textDefaultColorRect.Fill = new SolidColorBrush(Graphics.MediaColorFromDrawingColor(defaultTextColor));
            backgroundHighlightColorRect.Fill = new SolidColorBrush(Graphics.MediaColorFromDrawingColor(defaultBackgroundColor));
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            Drawing.Color textHighlightColor = Graphics.DrawingColorFromBrush(textHighlightColorRect.Fill);
            Drawing.Color textDefaultColor = Graphics.DrawingColorFromBrush(textDefaultColorRect.Fill);
            Drawing.Color backgroundHighlightColor = Graphics.DrawingColorFromBrush(backgroundHighlightColorRect.Fill);
            SettingsHandler(textHighlightColor, textDefaultColor, backgroundHighlightColor);
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void TextHighlightColorRect_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Color currentColor = (textHighlightColorRect.Fill as SolidColorBrush).Color;
            ColorDialog colorDialog = new ColorDialog();
            colorDialog.Color = Graphics.DrawingColorFromMediaColor(currentColor);
            colorDialog.FullOpen = true;
            if (colorDialog.ShowDialog() != Forms.DialogResult.Cancel)
            {
                textHighlightColorRect.Fill = Graphics.MediaBrushFromDrawingColor(colorDialog.Color);
            }
        }

        private void TextDefaultColorRect_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Color currentColor = (textDefaultColorRect.Fill as SolidColorBrush).Color;
            ColorDialog colorDialog = new ColorDialog();
            colorDialog.Color = Graphics.DrawingColorFromMediaColor(currentColor);
            colorDialog.FullOpen = true;
            if (colorDialog.ShowDialog() != Forms.DialogResult.Cancel)
            {
                textDefaultColorRect.Fill = Graphics.MediaBrushFromDrawingColor(colorDialog.Color);
            }
        }

        private void BackgroundHighlightColorRect_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Color currentColor = (backgroundHighlightColorRect.Fill as SolidColorBrush).Color;
            ColorDialog colorDialog = new ColorDialog();
            colorDialog.Color = Graphics.DrawingColorFromMediaColor(currentColor);
            colorDialog.FullOpen = true;
            if (colorDialog.ShowDialog() != Forms.DialogResult.Cancel)
            {
                backgroundHighlightColorRect.Fill = Graphics.MediaBrushFromDrawingColor(colorDialog.Color);
            }
        }
    }
}
