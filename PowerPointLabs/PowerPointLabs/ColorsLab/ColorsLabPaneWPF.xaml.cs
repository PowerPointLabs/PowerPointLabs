using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.DataSources;

using Color = System.Drawing.Color;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ColorsLab
{
    
    /// <summary>
    /// Interaction logic for TimerLabPaneWPF.xaml
    /// </summary>
    public partial class ColorsLabPaneWPF : UserControl
    {
        // To set color mode
        private enum MODE
        {
            FILL,
            LINE,
            FONT,
            NONE
        };

        private Brush _previousFill;
        private PowerPoint.ShapeRange _selectedShapes;
        private PowerPoint.TextRange _selectedText;

        // Data-bindings datasource
        ColorDataSource dataSource = new ColorDataSource();

        public ColorsLabPaneWPF()
        {
            // Set data context to data source for XAML to reference.
            DataContext = dataSource;

            // Do not remove. Default generated code.
            InitializeComponent();

            // Setup code
            SetupImageSources();
            SetupRectangleClickEvents();
            SetDefaultColor(Color.CornflowerBlue);
        }

        #region Setup Code

        /// <summary>
        /// Function that handles the setting up of all the images in the pane.
        /// </summary>
        private void SetupImageSources()
        {
            textColorIcon.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    Properties.Resources.TextColor_icon.GetHbitmap(),
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());

            lineColorIcon.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    Properties.Resources.LineColor_icon.GetHbitmap(),
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());

            fillColorIcon.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    Properties.Resources.FillColor_icon.GetHbitmap(),
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());

            brightnessIcon.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    Properties.Resources.Brightness_icon_25x25.GetHbitmap(),
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());

            saturationIcon.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    Properties.Resources.Saturation_icon_18x18.GetHbitmap(),
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());

            saveColorIcon.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    Properties.Resources.Save_icon.GetHbitmap(),
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());

            loadColorIcon.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    Properties.Resources.Load_icon.GetHbitmap(),
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());

            reloadColorIcon.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    Properties.Resources.Reload_icon.GetHbitmap(),
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());

            clearColorIcon.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    Properties.Resources.Clear_icon.GetHbitmap(),
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());
        }

        /// <summary>
        /// Handles setting up of the click events for all the color rectangles in the pane.
        /// We are setting up the click events in code because WPF Rectangles do not have an intrinsic OnClick event.
        /// </summary>
        private void SetupRectangleClickEvents()
        {
            selectedColorRectangle.MouseDown += SelectedColorRectangle_Click;
            monochromaticRectangleOne.MouseDown += MatchingColorsRectangle_Click;
            monochromaticRectangleTwo.MouseDown += MatchingColorsRectangle_Click;
            monochromaticRectangleThree.MouseDown += MatchingColorsRectangle_Click;
            monochromaticRectangleFour.MouseDown += MatchingColorsRectangle_Click;
            monochromaticRectangleFive.MouseDown += MatchingColorsRectangle_Click;
            monochromaticRectangleSix.MouseDown += MatchingColorsRectangle_Click;
            monochromaticRectangleSeven.MouseDown += MatchingColorsRectangle_Click;
            analogousLowerRectangle.MouseDown += MatchingColorsRectangle_Click;
            analogousMiddleRectangle.MouseDown += MatchingColorsRectangle_Click;
            analogousHigherRectangle.MouseDown += MatchingColorsRectangle_Click;
            complementaryLowerRectangle.MouseDown += MatchingColorsRectangle_Click;
            complementaryMiddleRectangle.MouseDown += MatchingColorsRectangle_Click;
            complementaryHigherRectangle.MouseDown += MatchingColorsRectangle_Click;
            triadicLowerRectangle.MouseDown += MatchingColorsRectangle_Click;
            triadicMiddleRectangle.MouseDown += MatchingColorsRectangle_Click;
            triadicHigherRectangle.MouseDown += MatchingColorsRectangle_Click;
            tetradicOneRectangle.MouseDown += MatchingColorsRectangle_Click;
            tetradicTwoRectangle.MouseDown += MatchingColorsRectangle_Click;
            tetradicThreeRectangle.MouseDown += MatchingColorsRectangle_Click;
            tetradicFourRectangle.MouseDown += MatchingColorsRectangle_Click;
        }

        /// <summary>
        /// Set default color upon startup.
        /// </summary>
        /// <param name="color"></param>
        private void SetDefaultColor(Color color)
        {
            dataSource.SelectedColor = color;
        }

        #endregion

        #region Event Handlers

        #region Button Handlers

        private void ApplyTextColorButton_Click(object sender, RoutedEventArgs e)
        {
            ColorSelectedShapesWithColor(dataSource.SelectedColor, MODE.FONT);
        }

        private void ApplyLineColorButton_Click(object sender, RoutedEventArgs e)
        {
            ColorSelectedShapesWithColor(dataSource.SelectedColor, MODE.LINE);
        }

        private void ApplyFillColorButton_Click(object sender, RoutedEventArgs e)
        {
            ColorSelectedShapesWithColor(dataSource.SelectedColor, MODE.FILL);
        }

        private void ApplyTextColorButton_MouseEnter(object sender, MouseEventArgs e)
        {

        }

        private void ApplyTextColorButton_MouseLeave(object sender, MouseEventArgs e)
        {

        }

        private void ApplyLineColorButton_MouseEnter(object sender, MouseEventArgs e)
        {

        }

        private void ApplyLineColorButton_MouseLeave(object sender, MouseEventArgs e)
        {

        }

        private void ApplyFillColorButton_MouseEnter(object sender, MouseEventArgs e)
        {

        }

        private void ApplyFillColorButton_MouseLeave(object sender, MouseEventArgs e)
        {

        }

        #endregion

        #region Slider Value Changed Handlers

        /// <summary>
        /// Updates selected color when brightness slider is moved.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BrightnessSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            double newBrightness = e.NewValue;
            HSLColor newColor = new HSLColor();
            newColor.Hue = dataSource.SelectedColor.Hue;
            newColor.Saturation = dataSource.SelectedColor.Saturation;
            newColor.Luminosity = newBrightness;
            dataSource.SelectedColor = newColor;
        }
        
        /// <summary>
        /// Updates selected color when saturation slider is moved.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SaturationSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            double newSaturation = e.NewValue;
            HSLColor newColor = new HSLColor();
            newColor.Hue = dataSource.SelectedColor.Hue;
            newColor.Saturation = newSaturation;
            newColor.Luminosity = dataSource.SelectedColor.Luminosity;
            dataSource.SelectedColor = newColor;
        }

        #endregion

        #region Color Rectangle Click Handlers

        /// <summary>
        /// Opens up a Windows.Forms ColorDialog upon click of the selectedColor rectangle.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SelectedColorRectangle_Click(object sender, MouseButtonEventArgs e)
        {
            System.Windows.Forms.ColorDialog colorPickerDialog = new System.Windows.Forms.ColorDialog();
            colorPickerDialog.FullOpen = true;

            // Sets the initial color select to the current selected color.
            colorPickerDialog.Color = dataSource.SelectedColor;

            // Update the selected color if the user clicks OK
            if (colorPickerDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                dataSource.SelectedColor = colorPickerDialog.Color;
            }
        }

        /// <summary>
        /// Updates the selected color in the data source when rectangle is clicked.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MatchingColorsRectangle_Click(object sender, MouseButtonEventArgs e)
        {
            System.Windows.Shapes.Rectangle rect = (System.Windows.Shapes.Rectangle)sender;
            System.Windows.Media.Color color = ((SolidColorBrush)rect.Fill).Color;
            Color selectedColor = Color.FromArgb(color.A, color.R, color.G, color.B);
            dataSource.SelectedColor = new HSLColor(selectedColor);
        }

        #endregion

        #endregion

        private void MonochromaticRectangleOne_MouseMove(object sender, MouseEventArgs e)
        {
            System.Windows.Shapes.Rectangle rect = (System.Windows.Shapes.Rectangle)sender;
            if (rect != null && e.LeftButton == MouseButtonState.Pressed)
            {
                DragDrop.DoDragDrop(rect, rect.Fill.ToString(), DragDropEffects.Copy);
            }
        }

        private void FavoriteRect_DragEnter(object sender, DragEventArgs e)
        {
            System.Windows.Shapes.Rectangle rect = (System.Windows.Shapes.Rectangle)sender;
            if (rect != null)
            {
                // Save the current Fill brush so that you can revert back to this value in DragLeave.
                _previousFill = rect.Fill;

                // If the DataObject contains string data, extract it.
                if (e.Data.GetDataPresent(DataFormats.StringFormat))
                {
                    string dataString = (string)e.Data.GetData(DataFormats.StringFormat);

                    // If the string can be converted into a Brush, convert it.
                    BrushConverter converter = new BrushConverter();
                    if (converter.IsValid(dataString))
                    {
                        System.Windows.Media.Brush newFill = (System.Windows.Media.Brush)converter.ConvertFromString(dataString);
                        rect.Fill = newFill;
                    }
                }
            }
        }

        private void FavoriteRect_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.None;

            // If the DataObject contains string data, extract it.
            if (e.Data.GetDataPresent(DataFormats.StringFormat))
            {
                string dataString = (string)e.Data.GetData(DataFormats.StringFormat);

                // If the string can be converted into a Brush, allow copying.
                BrushConverter converter = new BrushConverter();
                if (converter.IsValid(dataString))
                {
                    e.Effects = DragDropEffects.Copy | DragDropEffects.Move;
                }
            }
        }

        private void FavoriteRect_DragLeave(object sender, DragEventArgs e)
        {
            System.Windows.Shapes.Rectangle rect = (System.Windows.Shapes.Rectangle)sender;
            if (rect != null)
            {
                rect.Fill = _previousFill;
            }
        }

        private void FavoriteRect_Drop(object sender, DragEventArgs e)
        {
            System.Windows.Shapes.Rectangle rect = (System.Windows.Shapes.Rectangle)sender;
            if (rect != null)
            {
                // If the DataObject contains string data, extract it.
                if (e.Data.GetDataPresent(DataFormats.StringFormat))
                {
                    string dataString = (string)e.Data.GetData(DataFormats.StringFormat);

                    // If the string can be converted into a Brush, 
                    // convert it and apply it to the ellipse.
                    BrushConverter converter = new BrushConverter();
                    if (converter.IsValid(dataString))
                    {
                        System.Windows.Media.Brush newFill = (System.Windows.Media.Brush)converter.ConvertFromString(dataString);
                        rect.Fill = newFill;
                    }
                }
            }
        }

        private void ColorSelectedShapesWithColor(HSLColor selectedColor, MODE colorMode)
        {
            SelectShapes();
            if (_selectedShapes != null
                && this.GetCurrentSelection().Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape s in _selectedShapes)
                {
                    try
                    {
                        Byte r = ((Color)selectedColor).R;
                        Byte g = ((Color)selectedColor).G;
                        Byte b = ((Color)selectedColor).B;

                        int rgb = (b << 16) | (g << 8) | (r);
                        ColorShapeWithColor(s, rgb, colorMode);
                    }
                    catch (Exception)
                    {
                        RecreateCorruptedShape(s);
                    }
                }
            }
            if (_selectedText != null
                && this.GetCurrentSelection().Type == PowerPoint.PpSelectionType.ppSelectionText)
            {
                try
                {
                    Byte r = ((Color)selectedColor).R;
                    Byte g = ((Color)selectedColor).G;
                    Byte b = ((Color)selectedColor).B;

                    int rgb = (b << 16) | (g << 8) | (r);
                    ColorTextWithColor(_selectedText, rgb, colorMode);
                }
                catch (Exception)
                {
                }
            }
        }

        private void SelectShapes()
        {
            try
            {
                PowerPoint.Selection selection = this.GetCurrentSelection();
                if (selection == null)
                {
                    return;
                }

                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes &&
                    selection.HasChildShapeRange)
                {
                    _selectedShapes = selection.ChildShapeRange;
                }
                else if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    _selectedShapes = selection.ShapeRange;
                }
                else if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    _selectedText = selection.TextRange;
                }
                else
                {
                    _selectedShapes = null;
                    _selectedText = null;
                }
            }
            catch (Exception)
            {
                _selectedShapes = null;
                _selectedText = null;
            }
        }

        private void ColorTextWithColor(PowerPoint.TextRange text, int rgb, MODE mode)
        {
            PowerPoint.TextFrame frame = text.Parent as PowerPoint.TextFrame;
            PowerPoint.Shape selectedShape = frame.Parent as PowerPoint.Shape;
            if (mode != MODE.NONE)
            {
                ColorShapeWithColor(selectedShape, rgb, mode);
            }
        }

        private void ColorShapeWithColor(PowerPoint.Shape s, int rgb, MODE mode)
        {
            switch (mode)
            {
                case MODE.FILL:
                    s.Fill.ForeColor.RGB = rgb;
                    break;
                case MODE.LINE:
                    s.Line.ForeColor.RGB = rgb;
                    s.Line.Visible = MsoTriState.msoTrue;
                    break;
                case MODE.FONT:
                    ColorShapeFontWithColor(s, rgb);
                    break;
            }
        }

        private void ColorShapeFontWithColor(PowerPoint.Shape s, int rgb)
        {
            if (s.HasTextFrame == MsoTriState.msoTrue)
            {
                PowerPoint.Selection selection = this.GetCurrentSelection();
                if (selection == null)
                {
                    return;
                }

                if (selection.ShapeRange.HasTextFrame == MsoTriState.msoTrue)
                {
                    if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                    {
                        PowerPoint.TextRange selectedText = selection.TextRange.TrimText();
                        if (selectedText.Text != "" && selectedText != null)
                        {
                            selectedText.Font.Color.RGB = rgb;
                        }
                        else
                        {
                            s.TextFrame.TextRange.TrimText().Font.Color.RGB = rgb;
                        }
                    }
                    else if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        s.TextFrame.TextRange.TrimText().Font.Color.RGB = rgb;
                    }
                }
            }
        }

        private void RecreateCorruptedShape(PowerPoint.Shape s)
        {
            s.Copy();
            PowerPoint.Shape newShape = this.GetCurrentSlide().Shapes.Paste()[1];

            newShape.Select();

            newShape.Name = s.Name;
            newShape.Left = s.Left;
            newShape.Top = s.Top;
            while (newShape.ZOrderPosition > s.ZOrderPosition)
            {
                newShape.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendBackward);
            }
            s.Delete();
        }


    }
}
