using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Threading;
using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.DataSources;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Views;

using Color = System.Drawing.Color;
using Path = System.IO.Path;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ColorsLab
{

    /// <summary>
    /// Interaction logic for TimerLabPaneWPF.xaml
    /// </summary>
    public partial class ColorsLabPaneWPF : UserControl
    {
        #region Functional Test API

        public Point GetApplyTextButtonLocationAsPoint()
        {
            Point locationFromWindow = applyTextColorButton.TranslatePoint(new Point(0, 0), this);
            Point topLeftOfButton = PointToScreen(locationFromWindow);
            return new Point(
                topLeftOfButton.X + applyTextColorButton.ActualWidth / 2, 
                topLeftOfButton.Y + applyTextColorButton.ActualHeight / 2);
        }

        public Point GetApplyLineButtonLocationAsPoint()
        {
            Point locationFromWindow = applyLineColorButton.TranslatePoint(new Point(0, 0), this);
            Point topLeftOfButton = PointToScreen(locationFromWindow);
            return new Point(
                topLeftOfButton.X + applyLineColorButton.ActualWidth / 2,
                topLeftOfButton.Y + applyLineColorButton.ActualHeight / 2);
        }

        public Point GetApplyFillButtonLocationAsPoint()
        {
            Point locationFromWindow = applyFillColorButton.TranslatePoint(new Point(0, 0), this);
            Point topLeftOfButton = PointToScreen(locationFromWindow);
            return new Point(
                topLeftOfButton.X + applyFillColorButton.ActualWidth / 2,
                topLeftOfButton.Y + applyFillColorButton.ActualHeight / 2);
        }

        public Point GetMainColorRectangleLocationAsPoint()
        {
            Point locationFromWindow = selectedColorRectangle.TranslatePoint(new Point(0, 0), this);
            Point topLeftOfButton = PointToScreen(locationFromWindow);
            return new Point(
                topLeftOfButton.X + selectedColorRectangle.ActualWidth / 2,
                topLeftOfButton.Y + selectedColorRectangle.ActualHeight / 2);
        }

        public Point GetEyeDropperButtonLocationAsPoint()
        {
            Point locationFromWindow = eyeDropperButton.TranslatePoint(new Point(0, 0), this);
            Point topLeftOfButton = PointToScreen(locationFromWindow);
            return new Point(
                topLeftOfButton.X + eyeDropperButton.ActualWidth / 2,
                topLeftOfButton.Y + eyeDropperButton.ActualHeight / 2);
        }

        public IList<Color> GetFavoriteColorsPanelAsList()
        {
            IList<HSLColor> favoriteHslColors = dataSource.GetListOfFavoriteColors();
            IList<Color> favoriteColors = new List<Color>();
            foreach (HSLColor favoriteHslColor in favoriteHslColors)
            {
                favoriteColors.Add(favoriteHslColor);
            }
            return favoriteColors;
        }

        public void LoadFavoriteColorsFromPath(string filePath)
        {
            dataSource.LoadFavoriteColorsFromFile(filePath);
        }

        public IList<Color> GetRecentColorsPanelAsList()
        {
            IList<HSLColor> recentHslColors = dataSource.GetListOfRecentColors();
            IList<Color> recentColors = new List<Color>();
            foreach (HSLColor recentHslColor in recentHslColors)
            {
                recentColors.Add(recentHslColor);
            }
            return recentColors;
        }

        /// <summary>
        /// Clear the panel to all white color.
        /// </summary>
        public void EmptyRecentColorsPanel()
        {
            try
            {
                if (this.GetCurrentSlide() != null)
                {
                    dataSource.ClearRecentColors();
                }
            }
            catch (Exception e)
            {
                ErrorDialogBox.ShowDialog("Recent Colors Panel Reset Failed", e.Message, e);
            }
        }

        #endregion

        #region Private variables

        // To set color mode
        private enum MODE
        {
            FILL,
            LINE,
            FONT,
            MAIN,
            NONE
        };

        private MODE _eyedropperMode;
        private Color _previousFillColor;
        private Color _currentEyedroppedColor;
        private Color _currentSelectedColor;
        private PowerPoint.ShapeRange _selectedShapes;
        private PowerPoint.TextRange _selectedText;
        private bool _isEyedropperMode = false;
        private bool _shouldAllowDrag = false;

        // Data-bindings datasource
        ColorDataSource dataSource = new ColorDataSource();


        // Eyedropper-related
        private const float MAGNIFICATION_FACTOR = 2.5f;
        private Cursor eyeDropperCursor = new Cursor(new MemoryStream(Properties.Resources.EyeDropper));
        private Magnifier magnifier = new Magnifier(MAGNIFICATION_FACTOR);
        private System.Windows.Forms.Timer eyeDropperTimer = new System.Windows.Forms.Timer(new System.ComponentModel.Container());
        private const int CLICK_THRESHOLD = 2;
        private int timer1Ticks;

        // Saving color themes
        private string _defaultThemeColorDirectory = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), 
            "PowerPointLabs.defaultThemeColor.thm");
        private string _defaultRecentColorDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "PowerPointLabs.defaultRecentColor.thm");

        #endregion

        #region Constructor

        public ColorsLabPaneWPF()
        {
            // Set data context to data source for XAML to reference.
            DataContext = dataSource;

            // Do not remove. Default generated code.
            InitializeComponent();

            // Setup code
            SetupImageSources();
            SetupEyedropperTimer();
            SetDefaultColor(Color.CornflowerBlue);
            SetDefaultThemeColors();
            SetRecentColors();

            // Hook the mouse process if it has not
            PPExtraEventHelper.PPMouse.TryStartHook();
        }

        #endregion

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

            eyeDropperIcon.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    Properties.Resources.EyeDropper_icon.GetHbitmap(),
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

        /// <summary>
        /// Set default theme colors for favourite colors panel.
        /// </summary>
        private void SetDefaultThemeColors()
        {
            LoadDefaultThemePanel();
        }

        /// <summary>
        /// Load recent colors into the recent colors pane.
        /// </summary>
        private void SetRecentColors()
        {
            LoadRecentColorsPanel();
        }

        /// <summary>
        /// Setup the timer tick handler.
        /// </summary>
        private void SetupEyedropperTimer()
        {
            this.eyeDropperTimer.Tick += new System.EventHandler(this.Timer1_Tick);
        }

        #endregion

        #region Event Handlers
        
        #region ColorsLabPane Handlers

        /// <summary>
        /// This method handles the loaded ColorsLabPane.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ColorsLabPaneWPF_Loaded(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Tools.CustomTaskPane colorsLabPane = this.GetAddIn().GetActivePane(typeof(ColorsLabPane));
            if (colorsLabPane == null || !(colorsLabPane.Control is ColorsLabPane))
            {
                MessageBox.Show("Error: ColorsLabPane not opened.");
                return;
            }
            ColorsLabPane colorsLab = colorsLabPane.Control as ColorsLabPane;

            // Add handler for closing of ColorsLab
            colorsLab.HandleDestroyed += ColorsLab_Closing;
        }

        /// <summary>
        /// This handler is called when ColorsLab is destroyed, i.e. when PPTLabs closes.
        /// Current colors in the panel are saved when this happens.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ColorsLab_Closing(Object sender, EventArgs e)
        {
            SaveDefaultColorPaneThemeColors();
            SaveColorPaneRecentColors();
        }

        #endregion

        #region Button Handlers

        /// <summary>
        /// On mouse down, init eyedropper mode.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ApplyColorButton_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (this.GetCurrentSelection().Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                MessageBox.Show(ColorsLabText.ErrorNoSelection, ColorsLabText.ErrorDialogTitle);
                return;
            }

            CaptureMouse();
            SetEyedropperMode(((Button)sender).Name);
            BeginEyedropping();
            this.GetApplication().StartNewUndoEntry();
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

        /// <summary>
        /// Handles drag-and-drop functionality for color rects that can be dragged to favourite colors.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DraggableColorRectangle_MouseMove(object sender, MouseEventArgs e)
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
                try
                {
                    DragDrop.DoDragDrop(rect, rect.Fill.ToString(), DragDropEffects.Copy);
                } 
                catch (System.Runtime.InteropServices.COMException)
                {
                    // This exception occurs when user tries to drag the color rect to a textbox/shape on the slide.
                    // Due to lack of drag and drop support for some PowerPoint objects, exception will be thrown.
                    // When this is detected, to insert the data to the textbox instead.
                    // More info: https://social.msdn.microsoft.com/Forums/en-US/9925d6c7-e92f-40e7-9467-7b4e69174e9e/vsto-addin-gt-facing-problem-in-implementing-dragdrop-functionality-gt-need-help?forum=vsto
                    PowerPoint.Selection currSelection = this.GetCurrentSelection();
                    if (currSelection.Type == PowerPoint.PpSelectionType.ppSelectionShapes ||
                        currSelection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                    {
                        currSelection.TextRange2.Text = rect.Fill.ToString();
                    }
                }
                _shouldAllowDrag = false;
            }
        }

        /// <summary>
        /// This function prevents user to begin the drag from outside the rectangle.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DraggableColorRectangle_MouseLeave(object sender, MouseEventArgs e)
        {
            if (_isEyedropperMode)
            {
                return;
            }

            _shouldAllowDrag = false;

            System.Windows.Shapes.Rectangle rect = (System.Windows.Shapes.Rectangle)sender;

            if (rect.Name == "selectedColorRectangle")
            {
                rect.MouseUp -= SelectedColorRectangle_MouseUp;
            }
        }

        /// <summary>
        /// Adds a MouseUp listener to the sender object when it detects a MouseDown.
        /// The purpose of this is to simulate a click event on the Rectangle object.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ColorRectangle_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (_isEyedropperMode)
            {
                return;
            }

            // We remove the MouseUp event first before adding it to ensure that at anytime there's only
            // one listener for the MouseUp event.
            System.Windows.Shapes.Rectangle rect = (System.Windows.Shapes.Rectangle)sender;
            rect.MouseUp -= ColorRectangle_MouseUp;
            rect.MouseUp += ColorRectangle_MouseUp;
        }

        /// <summary>
        /// Change the selectedColor to the color of the matching color rect.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ColorRectangle_MouseUp(object sender, MouseButtonEventArgs e)
        {
            System.Windows.Shapes.Rectangle rect = (System.Windows.Shapes.Rectangle)sender;
            rect.MouseUp -= ColorRectangle_MouseUp;

            System.Windows.Media.Color color = ((SolidColorBrush)rect.Fill).Color;
            Color selectedColor = Color.FromArgb(color.A, color.R, color.G, color.B);
            dataSource.SelectedColor = new HSLColor(selectedColor);
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

            System.Windows.Media.Color prevMediaColor = ((SolidColorBrush)rect.Fill).Color;
            _previousFillColor = Color.FromArgb(prevMediaColor.A, prevMediaColor.R, prevMediaColor.G, prevMediaColor.B);

            if (rect != null)
            {
                // If the DataObject contains string data, extract it.
                if (e.Data.GetDataPresent(DataFormats.StringFormat))
                {
                    string dataString = (string)e.Data.GetData(DataFormats.StringFormat);

                    // If the string can be converted into a Color, 
                    // convert it and apply it to the rect.
                    ColorConverter converter = new ColorConverter();
                    if (converter.IsValid(dataString))
                    {
                        System.Windows.Media.Color mediaColor = (System.Windows.Media.Color)ColorConverter.ConvertFromString(dataString);
                        Color color = Color.FromArgb(mediaColor.A, mediaColor.R, mediaColor.G, mediaColor.B);

                        SetThemeColorRectangle(Grid.GetColumn(rect), color);
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
                SetThemeColorRectangle(Grid.GetColumn(rect), _previousFillColor);
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

                    // If the string can be converted into a Color, 
                    // convert it and apply it to the rect.
                    ColorConverter converter = new ColorConverter();
                    if (converter.IsValid(dataString))
                    {
                        System.Windows.Media.Color mediaColor = (System.Windows.Media.Color)ColorConverter.ConvertFromString(dataString);
                        Color color = Color.FromArgb(mediaColor.A, mediaColor.R, mediaColor.G, mediaColor.B);

                        SetThemeColorRectangle(Grid.GetColumn(rect), color);
                    }
                }
            }
        }

        #endregion

        #region Eye Dropper Event Handlers

        /// <summary>
        /// Gets the mouse position and pixel color value every tick of the timer.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Timer1_Tick(object sender, EventArgs e)
        {
            timer1Ticks++;

            System.Drawing.Point mousePos = System.Windows.Forms.Control.MousePosition;
            IntPtr deviceContext = PPExtraEventHelper.Native.GetDC(IntPtr.Zero);

            Color pickedColor = System.Drawing.ColorTranslator.FromWin32(PPExtraEventHelper.Native.GetPixel(deviceContext, mousePos.X, mousePos.Y));

            // If button has not been held long enough to register as a drag, then don't pick a color.
            if (timer1Ticks < CLICK_THRESHOLD)
            {
                return;
            }

            if (_eyedropperMode == MODE.MAIN)
            {
                ColorMainColorRect(pickedColor);
                _currentEyedroppedColor = pickedColor;
            }
            else
            {
                ColorSelectedShapesWithColor(pickedColor, _eyedropperMode);
                _currentSelectedColor = pickedColor;
            }
        }

        /// <summary>
        /// Handles the end of eye dropper mode.
        /// </summary>
        void LeftMouseButtonUpEventHandler()
        {
            PPExtraEventHelper.PPMouse.LeftButtonUp -= LeftMouseButtonUpEventHandler;
            magnifier.Hide();
            eyeDropperTimer.Stop();

            // Only change the actual datasource selectedColor at the end of eyedropping.
            if (_eyedropperMode == MODE.MAIN)
            {
                selectedColorRectangle.Opacity = 1;
                if (timer1Ticks > CLICK_THRESHOLD)
                {
                    dataSource.SelectedColor = _currentEyedroppedColor;
                }
            }

            // Update recent colors if color has been used
            if (_eyedropperMode == MODE.FILL || _eyedropperMode == MODE.FONT || _eyedropperMode == MODE.LINE)
            {
                if (timer1Ticks > CLICK_THRESHOLD)
                {
                    dataSource.AddColorToRecentColors(_currentSelectedColor);
                }
            }

            _isEyedropperMode = false;
            _eyedropperMode = MODE.NONE;
            Mouse.OverrideCursor = null;
            Mouse.Capture(null);

            if (timer1Ticks < CLICK_THRESHOLD)
            {
                Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Loaded, new Action(delegate
                {
                    MessageBox.Show("To use this button, click and drag to desired color.", ColorsLabText.ErrorDialogTitle);
                }));
            }
        }

        private void EyeDropperButton_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            CaptureMouse();
            SetEyedropperMode(((Button)sender).Name);
            BeginEyedropping();
            this.GetApplication().StartNewUndoEntry();
        }

        #endregion

        #region Favourite Colors Button Handlers

        /// <summary>
        /// Saves the current theme panel as a custom theme file.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SaveColorButton_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            saveFileDialog.DefaultExt = "thm";
            saveFileDialog.Filter = "PPTLabsThemes|*.thm";
            saveFileDialog.Title = "Save Theme";
            
            System.Windows.Forms.DialogResult result = saveFileDialog.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK &&
                dataSource.SaveFavoriteColorsInFile(saveFileDialog.FileName))
            {
                SaveDefaultColorPaneThemeColors();
            }
        }

        /// <summary>
        /// Loads a theme panel from an existing theme file.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LoadColorButton_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();
            openFileDialog.DefaultExt = "thm";
            openFileDialog.Filter = "PPTLabsTheme|*.thm";
            openFileDialog.Title = "Load Theme";

            System.Windows.Forms.DialogResult result = openFileDialog.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK &&
                dataSource.LoadFavoriteColorsFromFile(openFileDialog.FileName))
            {
                SaveDefaultColorPaneThemeColors();
            }
        }

        /// <summary>
        /// Reset to the last loaded theme panel.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ReloadColorButton_Click(object sender, EventArgs e)
        {
            ResetThemePanel();
        }

        /// <summary>
        /// Clears the current theme panel, i.e. set all favourite colors to white.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ClearColorButton_Click(object sender, RoutedEventArgs e)
        {
            EmptyThemePanel();
        }

        #endregion

        #region Context Menu Handlers

        private void Color_Information_Click(object sender, RoutedEventArgs e)
        {
            MenuItem menuItem = sender as MenuItem;
            if (menuItem == null)
            {
                return;
            }
            System.Windows.Shapes.Rectangle rect = ((ContextMenu)(menuItem.Parent)).PlacementTarget as System.Windows.Shapes.Rectangle;
            System.Windows.Media.Color color = ((SolidColorBrush)rect.Fill).Color;
            Color selectedColor = Color.FromArgb(color.A, color.R, color.G, color.B);
            ColorInformationDialog dialog = new ColorInformationDialog(selectedColor);
            dialog.Show();
        }

        private void Set_Main_Color_Click(object sender, RoutedEventArgs e)
        {
            MenuItem menuItem = sender as MenuItem;
            if (menuItem == null)
            {
                return;
            }
            System.Windows.Shapes.Rectangle rect = ((ContextMenu)(menuItem.Parent)).PlacementTarget as System.Windows.Shapes.Rectangle;
            System.Windows.Media.Color color = ((SolidColorBrush)rect.Fill).Color;
            Color selectedColor = Color.FromArgb(color.A, color.R, color.G, color.B);
            dataSource.SelectedColor = new HSLColor(selectedColor);
        }

        private void Add_Favorite_Click(object sender, RoutedEventArgs e)
        {
            MenuItem menuItem = sender as MenuItem;
            if (menuItem == null)
            {
                return;
            }
            System.Windows.Shapes.Rectangle rect = ((ContextMenu)(menuItem.Parent)).PlacementTarget as System.Windows.Shapes.Rectangle;
            System.Windows.Media.Color color = ((SolidColorBrush)rect.Fill).Color;
            HSLColor clickedColor = Color.FromArgb(color.A, color.R, color.G, color.B);
            dataSource.AddColorToFavorites(clickedColor);
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

        #region Eye Dropper

        /// <summary>
        /// Sets the eyedropper mode given the name of the rectangle.
        /// </summary>
        /// <param name="rectName"></param>
        private void SetEyedropperMode(string rectName)
        {
            switch (rectName)
            {
                case "applyTextColorButton":
                    _eyedropperMode = MODE.FONT;
                    break;
                case "applyLineColorButton":
                    _eyedropperMode = MODE.LINE;
                    break;
                case "applyFillColorButton":
                    _eyedropperMode = MODE.FILL;
                    break;
                case "eyeDropperButton":
                    _eyedropperMode = MODE.MAIN;
                    break;
                default:
                    _eyedropperMode = MODE.NONE;
                    break;
            }
        }

        /// <summary>
        /// Show magnifier and begin eye dropping.
        /// </summary>
        private void BeginEyedropping()
        {
            _isEyedropperMode = true;
            timer1Ticks = 0;
            eyeDropperTimer.Start();
            Mouse.OverrideCursor = eyeDropperCursor;
            PPExtraEventHelper.PPMouse.LeftButtonUp += LeftMouseButtonUpEventHandler;
            magnifier.Show();

            if (_eyedropperMode == MODE.MAIN)
            {
                eyeDropperPreviewRectangle.Fill = selectedColorRectangle.Fill;
                selectedColorRectangle.Opacity = 0;
            }
        }

        private void ColorMainColorRect(Color color)
        {
            eyeDropperPreviewRectangle.Fill = 
                new SolidColorBrush(System.Windows.Media.Color.FromArgb(color.A, color.R, color.G, color.B));
        }

        #endregion

        #region Favourite Colors

        /// <summary>
        /// Load default panel from default file, or clear the panel if unsuccessful.
        /// </summary>
        private void LoadDefaultThemePanel()
        {
            bool isSuccessful = dataSource.LoadFavoriteColorsFromFile(_defaultThemeColorDirectory);
            if (!isSuccessful)
            {
                EmptyThemePanel();
            }
        }

        /// <summary>
        /// Reset panel to the last loaded theme panel.
        /// </summary>
        private void ResetThemePanel()
        {
            try
            {
                LoadDefaultThemePanel();
            }
            catch (Exception e)
            {
                ErrorDialogBox.ShowDialog("Theme Panel Reset Failed", e.Message, e);
            }
        }

        /// <summary>
        /// Clear the panel to all white color.
        /// </summary>
        private void EmptyThemePanel()
        {
            try
            {
                if (this.GetCurrentSlide() != null)
                {
                    dataSource.ClearFavoriteColors();
                }
            }
            catch (Exception e)
            {
                ErrorDialogBox.ShowDialog("Theme Panel Reset Failed", e.Message, e);
            }
        }

        /// <summary>
        /// Set the color rect given the name of the rect, and the color to set.
        /// </summary>
        /// <param name="name"></param>
        /// <param name="color"></param>
        private void SetThemeColorRectangle(int column, Color color)
        {
            dataSource.SetFavoriteColor(column, color);
        }

        /// <summary>
        /// Save current panel as default theme color.
        /// </summary>
        private void SaveDefaultColorPaneThemeColors()
        {
            dataSource.SaveFavoriteColorsInFile(_defaultThemeColorDirectory);
        }

        #endregion

        #region Recent Colors
        
        /// <summary>
        /// Save current recent colors panel to file.
        /// </summary>
        private void SaveColorPaneRecentColors()
        {
            dataSource.SaveRecentColorsInFile(_defaultRecentColorDirectory);
        }

        /// <summary>
        /// Load recent panel from file, or clear the panel if unsuccessful.
        /// </summary>
        private void LoadRecentColorsPanel()
        {
            bool isSuccessful = dataSource.LoadRecentColorsFromFile(_defaultRecentColorDirectory);
            if (!isSuccessful)
            {
                EmptyRecentColorsPanel();
            }
        }


        #endregion

    }
}
