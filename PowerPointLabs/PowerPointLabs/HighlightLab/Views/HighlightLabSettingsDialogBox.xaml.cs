using System.Windows;
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
        public delegate void DialogConfirmedDelegate(Drawing.Color highlightColor, 
                                                    Drawing.Color defaultColor, 
                                                    Drawing.Color backgroundColor);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }

        public HighlightLabSettingsDialogBox()
        {
            InitializeComponent();
        }

        public HighlightLabSettingsDialogBox(Drawing.Color defaultHighlightColor, 
                                            Drawing.Color defaultTextColor, 
                                            Drawing.Color defaultBackgroundColor)
            : this()
        {
            textHighlightColorRect.Fill = new SolidColorBrush(GraphicsUtil.MediaColorFromDrawingColor(defaultHighlightColor));
            textDefaultColorRect.Fill = new SolidColorBrush(GraphicsUtil.MediaColorFromDrawingColor(defaultTextColor));
            backgroundHighlightColorRect.Fill = new SolidColorBrush(GraphicsUtil.MediaColorFromDrawingColor(defaultBackgroundColor));
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            Drawing.Color textHighlightColor = GraphicsUtil.DrawingColorFromBrush(textHighlightColorRect.Fill);
            Drawing.Color textDefaultColor = GraphicsUtil.DrawingColorFromBrush(textDefaultColorRect.Fill);
            Drawing.Color backgroundHighlightColor = GraphicsUtil.DrawingColorFromBrush(backgroundHighlightColorRect.Fill);
            DialogConfirmedHandler(textHighlightColor, textDefaultColor, backgroundHighlightColor);
            Close();
        }

        private void TextHighlightColorRect_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Color currentColor = (textHighlightColorRect.Fill as SolidColorBrush).Color;
            ColorDialog colorDialog = new ColorDialog();
            colorDialog.Color = GraphicsUtil.DrawingColorFromMediaColor(currentColor);
            colorDialog.FullOpen = true;
            if (colorDialog.ShowDialog() != Forms.DialogResult.Cancel)
            {
                textHighlightColorRect.Fill = GraphicsUtil.MediaBrushFromDrawingColor(colorDialog.Color);
            }
        }

        private void TextDefaultColorRect_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Color currentColor = (textDefaultColorRect.Fill as SolidColorBrush).Color;
            ColorDialog colorDialog = new ColorDialog();
            colorDialog.Color = GraphicsUtil.DrawingColorFromMediaColor(currentColor);
            colorDialog.FullOpen = true;
            if (colorDialog.ShowDialog() != Forms.DialogResult.Cancel)
            {
                textDefaultColorRect.Fill = GraphicsUtil.MediaBrushFromDrawingColor(colorDialog.Color);
            }
        }

        private void BackgroundHighlightColorRect_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Color currentColor = (backgroundHighlightColorRect.Fill as SolidColorBrush).Color;
            ColorDialog colorDialog = new ColorDialog();
            colorDialog.Color = GraphicsUtil.DrawingColorFromMediaColor(currentColor);
            colorDialog.FullOpen = true;
            if (colorDialog.ShowDialog() != Forms.DialogResult.Cancel)
            {
                backgroundHighlightColorRect.Fill = GraphicsUtil.MediaBrushFromDrawingColor(colorDialog.Color);
            }
        }
    }
}
