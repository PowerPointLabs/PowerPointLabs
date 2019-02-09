using System;
using System.Drawing;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;

using PowerPointLabs.DataSources;

namespace PowerPointLabs.ColorsLab
{
    /// <summary>
    /// Interaction logic for TimerLabPaneWPF.xaml
    /// </summary>
    public partial class ColorsLabPaneWPF : UserControl
    {
        
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

        #region Button Click Handlers

        private void ApplyTextColorButton_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Dummy code, to complete
            dataSource.SelectedColor = Color.Red;
        }

        private void ApplyLineColorButton_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Dummy code, to complete
            dataSource.SelectedColor = Color.Yellow;
        }

        private void ApplyFillColorButton_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Dummy code, to complete
            dataSource.SelectedColor = Color.Gold;
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
            System.Windows.Media.Color color = ((System.Windows.Media.SolidColorBrush)rect.Fill).Color;
            Color selectedColor = Color.FromArgb(color.A, color.R, color.G, color.B);
            dataSource.SelectedColor = new HSLColor(selectedColor);
        }

        #endregion

        #endregion

    }
}
