using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ColorPicker;
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
        private Color _previousColor;
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
            // TOOD: MouseDown is NOT ideal here. We want to select the rect's color
            // only if it's a click, not just a MouseDown or MouseUp. However, WPF Rects 
            // don't have a click event. Need to figure out a way here. This affects functionality
            // of drag and drop also, because the rects changes colour upon a click and drag and the 
            // wrong color is dragged to the favourites panel.

            // TODO: Acutally all these can be moved to XAML. It's better that way.

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
            _previousColor = Color.Empty;
        }

        private void ApplyLineColorButton_Click(object sender, RoutedEventArgs e)
        {
            ColorSelectedShapesWithColor(dataSource.SelectedColor, MODE.LINE);
            _previousColor = Color.Empty;
        }

        private void ApplyFillColorButton_Click(object sender, RoutedEventArgs e)
        {
            ColorSelectedShapesWithColor(dataSource.SelectedColor, MODE.FILL);
            _previousColor = Color.Empty;
        }

        private void ApplyTextColorButton_MouseEnter(object sender, MouseEventArgs e)
        {
            _previousColor = GetSelectedShapeColor(MODE.FONT);
            ColorSelectedShapesWithColor(dataSource.SelectedColor, MODE.FONT);
        }

        private void ApplyTextColorButton_MouseLeave(object sender, MouseEventArgs e)
        {
            if (_previousColor == Color.Empty)
            {
                return;
            }

            ColorSelectedShapesWithColor(_previousColor, MODE.FONT);
            _previousColor = Color.Empty;
        }

        private void ApplyLineColorButton_MouseEnter(object sender, MouseEventArgs e)
        {
            _previousColor = GetSelectedShapeColor(MODE.LINE);
            ColorSelectedShapesWithColor(dataSource.SelectedColor, MODE.LINE);
        }

        private void ApplyLineColorButton_MouseLeave(object sender, MouseEventArgs e)
        {
            if (_previousColor == Color.Empty)
            {
                return;
            }

            ColorSelectedShapesWithColor(_previousColor, MODE.LINE);
            _previousColor = Color.Empty;
        }

        private void ApplyFillColorButton_MouseEnter(object sender, MouseEventArgs e)
        {
            _previousColor = GetSelectedShapeColor(MODE.FILL);
            ColorSelectedShapesWithColor(dataSource.SelectedColor, MODE.FILL);
        }

        private void ApplyFillColorButton_MouseLeave(object sender, MouseEventArgs e)
        {
            if (_previousColor == Color.Empty)
            {
                return;
            }

            ColorSelectedShapesWithColor(_previousColor, MODE.FILL);
            _previousColor = Color.Empty;
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

        #region Color Rectangle Handlers

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

        /// <summary>
        /// Handles drag and drop functionality for matching colors rectangles.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MatchingColorsRectangle_MouseMove(object sender, MouseEventArgs e)
        {
            System.Windows.Shapes.Rectangle rect = (System.Windows.Shapes.Rectangle)sender;
            if (rect != null && e.LeftButton == MouseButtonState.Pressed)
            {
                DragDrop.DoDragDrop(rect, rect.Fill.ToString(), DragDropEffects.Copy);
            }
        }

        /// <summary>
        /// Handles drag and drop functionality for favourtie color rectangles.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        /// <summary>
        /// Handles drag and drop functionality for favourtie color rectangles.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        /// <summary>
        /// Handles drag and drop functionality for favourtie color rectangles.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FavoriteRect_DragLeave(object sender, DragEventArgs e)
        {
            System.Windows.Shapes.Rectangle rect = (System.Windows.Shapes.Rectangle)sender;
            if (rect != null)
            {
                rect.Fill = _previousFill;
            }
        }

        /// <summary>
        /// Handles drag and drop functionality for favourtie color rectangles.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
                    // convert it and apply it to the rect.
                    BrushConverter converter = new BrushConverter();
                    if (converter.IsValid(dataString))
                    {
                        Brush newFill = (Brush)converter.ConvertFromString(dataString);
                        rect.Fill = newFill;
                    }
                }
            }
        }

        #endregion

        #endregion

        #region Helpers

        #region Apply Colors (Text, Fill, Line)

        /// <summary>
        /// Color selected shapes with selected color, in the given mode.
        /// </summary>
        /// <param name="selectedColor"></param>
        /// <param name="colorMode"></param>
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

        /// <summary>
        /// Retrieves selected shapes or text.
        /// </summary>
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

        /// <summary>
        /// Colors specified text range with given color.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="rgb"></param>
        /// <param name="mode"></param>
        private void ColorTextWithColor(PowerPoint.TextRange text, int rgb, MODE mode)
        {
            PowerPoint.TextFrame frame = text.Parent as PowerPoint.TextFrame;
            PowerPoint.Shape selectedShape = frame.Parent as PowerPoint.Shape;
            if (mode != MODE.NONE)
            {
                ColorShapeWithColor(selectedShape, rgb, mode);
            }
        }

        /// <summary>
        /// Colors specified shape with color, in the given mode.
        /// </summary>
        /// <param name="s"></param>
        /// <param name="rgb"></param>
        /// <param name="mode"></param>
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

        /// <summary>
        /// Colors specified shape with color.
        /// </summary>
        /// <param name="s"></param>
        /// <param name="rgb"></param>
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

        /// <summary>
        /// Recreates any specified corrupted shape.
        /// </summary>
        /// <param name="s"></param>
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

        /// <summary>
        /// Retrieves color of the selected shape(s).
        /// </summary>
        /// <param name="mode"></param>
        /// <returns></returns>
        private Color GetSelectedShapeColor(MODE mode)
        {
            SelectShapes();
            if (_selectedShapes == null && _selectedText == null)
            {
                return dataSource.SelectedColor;
            }

            if (this.GetCurrentSelection().Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return GetSelectedShapeColor(_selectedShapes, mode);
            }
            if (this.GetCurrentSelection().Type == PowerPoint.PpSelectionType.ppSelectionText)
            {
                PowerPoint.TextFrame frame = _selectedText.Parent as PowerPoint.TextFrame;
                PowerPoint.Shape selectedShape = frame.Parent as PowerPoint.Shape;
                return GetSelectedShapeColor(selectedShape, mode);
            }

            return dataSource.SelectedColor;
        }

        /// <summary>
        /// Retrieves color of the selected shapeRange, 
        /// returning Black if shapeRange contains shapes with different colors.
        /// </summary>
        /// <param name="selectedShapes"></param>
        /// <param name="mode"></param>
        /// <returns></returns>
        private Color GetSelectedShapeColor(PowerPoint.ShapeRange selectedShapes, MODE mode)
        {
            Color colorToReturn = Color.Empty;
            foreach (object selectedShape in selectedShapes)
            {
                Color color = GetSelectedShapeColor(selectedShape as PowerPoint.Shape, mode);
                if (colorToReturn.Equals(Color.Empty))
                {
                    colorToReturn = color;
                }
                else if (!colorToReturn.Equals(color))
                {
                    return Color.Black;
                }
            }
            return colorToReturn;
        }

        /// <summary>
        /// Retrieves color of the selected shape.
        /// </summary>
        /// <param name="selectedShape"></param>
        /// <param name="mode"></param>
        /// <returns></returns>
        private Color GetSelectedShapeColor(PowerPoint.Shape selectedShape, MODE mode)
        {
            switch (mode)
            {
                case MODE.FILL:
                    return Color.FromArgb(ColorHelper.ReverseRGBToArgb(selectedShape.Fill.ForeColor.RGB));
                case MODE.LINE:
                    return Color.FromArgb(ColorHelper.ReverseRGBToArgb(selectedShape.Line.ForeColor.RGB));
                case MODE.FONT:
                    if (selectedShape.HasTextFrame == MsoTriState.msoTrue
                        && this.GetApplication().ActiveWindow.Selection.ShapeRange.HasTextFrame
                        == MsoTriState.msoTrue)
                    {
                        if (this.GetCurrentSelection().Type == PowerPoint.PpSelectionType.ppSelectionText)
                        {
                            PowerPoint.TextRange selectedText
                                = this.GetApplication().ActiveWindow.Selection.TextRange.TrimText();
                            if (selectedText != null && selectedText.Text != "")
                            {
                                return Color.FromArgb(ColorHelper.ReverseRGBToArgb(selectedText.Font.Color.RGB));
                            }
                            else
                            {
                                return
                                Color.FromArgb(
                                    ColorHelper.ReverseRGBToArgb(selectedShape.TextFrame.TextRange.Font.Color.RGB));
                            }
                        }
                        else if (this.GetCurrentSelection().Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                        {
                            return
                                Color.FromArgb(
                                    ColorHelper.ReverseRGBToArgb(selectedShape.TextFrame.TextRange.Font.Color.RGB));
                        }
                    }
                    break;
            }
            return dataSource.SelectedColor;
        }

        #endregion

        #endregion
    }
}
