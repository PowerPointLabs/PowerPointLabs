using System.Windows.Forms;
using System.Windows.Media;

using PowerPointLabs.Utils;

namespace PowerPointLabs.Views
{
    /// <summary>
    /// Interaction logic for HighlightLabSettingsDialogBox.xaml
    /// </summary>
    public partial class HighlightLabSettingsDialogBox
    {
        public delegate void UpdateSettingsDelegate(System.Drawing.Color highlightColor, 
                                                    System.Drawing.Color defaultColor, 
                                                    System.Drawing.Color backgroundColor);
        public UpdateSettingsDelegate SettingsHandler { get; set; }

        public HighlightLabSettingsDialogBox()
        {
            InitializeComponent();
        }

        public HighlightLabSettingsDialogBox(System.Drawing.Color defaultHighlightColor, 
                                            System.Drawing.Color defaultTextColor, 
                                            System.Drawing.Color defaultBackgroundColor)
            : this()
        {
            textHighlightColorRect.Fill = new SolidColorBrush(Graphics.MediaColorFromDrawingColor(defaultHighlightColor));
            textDefaultColorRect.Fill = new SolidColorBrush(Graphics.MediaColorFromDrawingColor(defaultTextColor));
            backgroundHighlightColorRect.Fill = new SolidColorBrush(Graphics.MediaColorFromDrawingColor(defaultBackgroundColor));
        }

        private void OkButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            System.Drawing.Color textHighlightColor = Graphics.DrawingColorFromBrush(textHighlightColorRect.Fill);
            System.Drawing.Color textDefaultColor = Graphics.DrawingColorFromBrush(textDefaultColorRect.Fill);
            System.Drawing.Color backgroundHighlightColor = Graphics.DrawingColorFromBrush(backgroundHighlightColorRect.Fill);
            SettingsHandler(textHighlightColor, textDefaultColor, backgroundHighlightColor);
            Close();
        }

        private void CancelButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Close();
        }

        private void TextHighlightColorRect_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            Color currentColor = (textHighlightColorRect.Fill as SolidColorBrush).Color;
            ColorDialog colorDialog = new ColorDialog();
            colorDialog.Color = Graphics.DrawingColorFromMediaColor(currentColor);
            colorDialog.FullOpen = true;
            if (colorDialog.ShowDialog() != System.Windows.Forms.DialogResult.Cancel)
            {
                textHighlightColorRect.Fill = Graphics.MediaBrushFromDrawingColor(colorDialog.Color);
            }
        }

        private void TextDefaultColorRect_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            Color currentColor = (textDefaultColorRect.Fill as SolidColorBrush).Color;
            ColorDialog colorDialog = new ColorDialog();
            colorDialog.Color = Graphics.DrawingColorFromMediaColor(currentColor);
            colorDialog.FullOpen = true;
            if (colorDialog.ShowDialog() != System.Windows.Forms.DialogResult.Cancel)
            {
                textDefaultColorRect.Fill = Graphics.MediaBrushFromDrawingColor(colorDialog.Color);
            }
        }

        private void BackgroundHighlightColorRect_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            Color currentColor = (backgroundHighlightColorRect.Fill as SolidColorBrush).Color;
            ColorDialog colorDialog = new ColorDialog();
            colorDialog.Color = Graphics.DrawingColorFromMediaColor(currentColor);
            colorDialog.FullOpen = true;
            if (colorDialog.ShowDialog() != System.Windows.Forms.DialogResult.Cancel)
            {
                backgroundHighlightColorRect.Fill = Graphics.MediaBrushFromDrawingColor(colorDialog.Color);
            }
        }
    }
}
