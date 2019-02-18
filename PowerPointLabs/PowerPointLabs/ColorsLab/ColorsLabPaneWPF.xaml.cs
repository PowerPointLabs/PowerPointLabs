﻿using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ColorPicker;
using PowerPointLabs.DataSources;
using PowerPointLabs.TextCollection;
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
        private bool _isEyedropperMode = false;
        private MODE _eyedropperMode;
        private bool _shouldAllowDrag = false;

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
            SetDefaultColor(Color.CornflowerBlue);

            this.timer1.Tick += new System.EventHandler(this.Timer1_Tick);

            // Hook the mouse process if it has not
            PPExtraEventHelper.PPMouse.TryStartHook();
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
            MessageBox.Show("hi");
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
        /// Add MouseUp event to rectangle to simulate a click event.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SelectedColorRectangle_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (_isEyedropperMode)
            {
                return;
            }

            // We remove the MouseUp event first before adding it to ensure that at anytime there's only
            // one listener for the MouseUp event.
            System.Windows.Shapes.Rectangle rect = (System.Windows.Shapes.Rectangle)sender;
            rect.MouseUp -= SelectedColorRectangle_MouseUp;
            rect.MouseUp += SelectedColorRectangle_MouseUp;
        }

        /// <summary>
        /// Opens up a Windows.Forms ColorDialog upon click of the selectedColor rectangle.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SelectedColorRectangle_MouseUp(object sender, MouseButtonEventArgs e)
        {
            System.Windows.Shapes.Rectangle rect = (System.Windows.Shapes.Rectangle)sender;
            rect.MouseUp -= SelectedColorRectangle_MouseUp;

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

        private void SelectedColorRectangle_MouseMove(object sender, MouseEventArgs e)
        {
            if (_isEyedropperMode)
            {
                return;
            }


            System.Windows.Shapes.Rectangle rect = (System.Windows.Shapes.Rectangle)sender;

            if (rect != null && e.LeftButton == MouseButtonState.Released)
            {
                _shouldAllowDrag = true;
            }

            if (rect != null && e.LeftButton == MouseButtonState.Pressed && _shouldAllowDrag)
            {
                DragDrop.DoDragDrop(rect, rect.Fill.ToString(), DragDropEffects.Copy);
                _shouldAllowDrag = false;
            }
        }

        private void SelectedColorRectangle_MouseLeave(object sender, MouseEventArgs e)
        {
            if (_isEyedropperMode)
            {
                return;
            }

            _shouldAllowDrag = false;

            System.Windows.Shapes.Rectangle rect = (System.Windows.Shapes.Rectangle)sender;
            rect.MouseUp -= SelectedColorRectangle_MouseUp;
        }

        /// <summary>
        /// Adds a MouseUp listener to the sender object when it detects a MouseDown.
        /// The purpose of this is to simulate a click event on the Rectangle object.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MatchingColorsRectangle_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (_isEyedropperMode)
            {
                return;
            }

            // We remove the MouseUp event first before adding it to ensure that at anytime there's only
            // one listener for the MouseUp event.
            System.Windows.Shapes.Rectangle rect = (System.Windows.Shapes.Rectangle)sender;
            rect.MouseUp -= MachingColorsRectangle_MouseUp;
            rect.MouseUp += MachingColorsRectangle_MouseUp;
        }

        /// <summary>
        /// Change the selectedColor to the color of the matching color rect.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MachingColorsRectangle_MouseUp(object sender, MouseButtonEventArgs e)
        {
            System.Windows.Shapes.Rectangle rect = (System.Windows.Shapes.Rectangle)sender;
            rect.MouseUp -= MachingColorsRectangle_MouseUp;

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
            if (_isEyedropperMode)
            {
                return;
            }

            System.Windows.Shapes.Rectangle rect = (System.Windows.Shapes.Rectangle)sender;

            if (rect != null && e.LeftButton == MouseButtonState.Released)
            {
                _shouldAllowDrag = true;
            }

            if (rect != null && e.LeftButton == MouseButtonState.Pressed && _shouldAllowDrag)
            {
                DragDrop.DoDragDrop(rect, rect.Fill.ToString(), DragDropEffects.Copy);
                _shouldAllowDrag = false;
            }
        }

        /// <summary>
        /// Handles drag and drop functionality for matching colors rectangles.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MatchingColorsRectangle_MouseLeave(object sender, MouseEventArgs e)
        {
            if (_isEyedropperMode)
            {
                return;
            }

            _shouldAllowDrag = false;
  
        }

        /// <summary>
        /// Handles drag and drop functionality for favourtie color rectangles.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FavoriteRect_DragEnter(object sender, DragEventArgs e)
        {
            if (_isEyedropperMode)
            {
                return;
            }

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
            if (_isEyedropperMode)
            {
                return;
            }

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
            if (_isEyedropperMode)
            {
                return;
            }

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
            if (_isEyedropperMode)
            {
                return;
            }

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

        private const float MAGNIFICATION_FACTOR = 2.5f;
        private Cursor eyeDropperCursor = new Cursor(new MemoryStream(Properties.Resources.EyeDropper));
        private Magnifier magnifier = new Magnifier(MAGNIFICATION_FACTOR);
        private System.Windows.Forms.Timer timer1 = new System.Windows.Forms.Timer(new System.ComponentModel.Container());
        private const int CLICK_THRESHOLD = 2;
        private int timer1Ticks;

        private void BeginEyedropping()
        {
            _isEyedropperMode = true;
            timer1Ticks = 0;
            timer1.Start();
            Mouse.OverrideCursor = eyeDropperCursor;
            PPExtraEventHelper.PPMouse.LeftButtonUp += LeftMouseButtonUpEventHandler;
            magnifier.Show();
        }

        private void Timer1_Tick(object sender, EventArgs e)
        {
            timer1Ticks++;

            System.Drawing.Point mousePos = System.Windows.Forms.Control.MousePosition;
            IntPtr deviceContext = PPExtraEventHelper.Native.GetDC(IntPtr.Zero);
            
            Color _pickedColor = System.Drawing.ColorTranslator.FromWin32(PPExtraEventHelper.Native.GetPixel(deviceContext, mousePos.X, mousePos.Y));
            ColorSelectedShapesWithColor(_pickedColor, _eyedropperMode);
        }

        void LeftMouseButtonUpEventHandler()
        {
            PPExtraEventHelper.PPMouse.LeftButtonUp -= LeftMouseButtonUpEventHandler;
            magnifier.Hide();
            timer1.Stop();

            // A click is detected, prompt user to drag.
            if (timer1Ticks < CLICK_THRESHOLD)
            {
                MessageBox.Show("Drag pls", ColorsLabText.ErrorDialogTitle);
            }

            _isEyedropperMode = false;
            _eyedropperMode = MODE.NONE;
            Mouse.OverrideCursor = null;
            ReleaseMouseCapture();
        }

        private void ApplyTextColorButton_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (this.GetCurrentSelection().Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                MessageBox.Show(ColorsLabText.ErrorNoSelection, ColorsLabText.ErrorDialogTitle);
                return;
            }

            CaptureMouse();
            _eyedropperMode = MODE.FONT;
            BeginEyedropping();
            this.GetApplication().StartNewUndoEntry();
        }

        private void ApplyLineColorButton_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (this.GetCurrentSelection().Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                MessageBox.Show(ColorsLabText.ErrorNoSelection, ColorsLabText.ErrorDialogTitle);
                return;
            }

            CaptureMouse();
            _eyedropperMode = MODE.LINE;
            BeginEyedropping();
            this.GetApplication().StartNewUndoEntry();
        }

        private void ApplyFillColorButton_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (this.GetCurrentSelection().Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                MessageBox.Show(ColorsLabText.ErrorNoSelection, ColorsLabText.ErrorDialogTitle);
                return;
            }

            CaptureMouse();
            _eyedropperMode = MODE.FILL;
            BeginEyedropping();
            this.GetApplication().StartNewUndoEntry();
        }

  
    }
}
