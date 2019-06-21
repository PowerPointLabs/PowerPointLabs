using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;

using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Windows;
using Drawing = System.Drawing;

using Forms = System.Windows.Forms;
using DrawingColor = System.Drawing.Color;

namespace PowerPointLabs.HighlightLab.Views
{
    /// <summary>
    /// Interaction logic for HighlightLabSettingsDialogBox.xaml
    /// </summary>
    public partial class HighlightLabSettingsDialogBox
    {
        public delegate void DialogConfirmedDelegate(DrawingColor highlightColor,
                                                    DrawingColor defaultColor,
                                                    DrawingColor backgroundColor);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }

        public HighlightLabSettingsDialogBox()
        {
            InitializeComponent();
        }

        public HighlightLabSettingsDialogBox(DrawingColor defaultHighlightColor,
                                            DrawingColor defaultTextColor,
                                            DrawingColor defaultBackgroundColor)
            : this()
        {
            textHighlightColorRect.Fill = new SolidColorBrush(GraphicsUtil.MediaColorFromDrawingColor(defaultHighlightColor));
            textDefaultColorRect.Fill = new SolidColorBrush(GraphicsUtil.MediaColorFromDrawingColor(defaultTextColor));
            backgroundHighlightColorRect.Fill = new SolidColorBrush(GraphicsUtil.MediaColorFromDrawingColor(defaultBackgroundColor));
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DrawingColor textHighlightColor = GraphicsUtil.DrawingColorFromBrush(textHighlightColorRect.Fill);
            DrawingColor textDefaultColor = GraphicsUtil.DrawingColorFromBrush(textDefaultColorRect.Fill);
            DrawingColor backgroundHighlightColor = GraphicsUtil.DrawingColorFromBrush(backgroundHighlightColorRect.Fill);
            DialogConfirmedHandler(textHighlightColor, textDefaultColor, backgroundHighlightColor);
            Close();
        }

        private void TextHighlightColorRect_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Color currentColor = (textHighlightColorRect.Fill as SolidColorBrush).Color;
            DrawingColor? resultColor = ColorDialogUtil.RequestForColor(currentColor);
            if (resultColor.HasValue)
            {
                textHighlightColorRect.Fill = GraphicsUtil.MediaBrushFromDrawingColor(resultColor.Value);
            }
        }

        private void TextDefaultColorRect_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Color currentColor = (textDefaultColorRect.Fill as SolidColorBrush).Color;
            DrawingColor? resultColor = ColorDialogUtil.RequestForColor(currentColor);
            if (resultColor.HasValue)
            {
                textDefaultColorRect.Fill = GraphicsUtil.MediaBrushFromDrawingColor(resultColor.Value);
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
