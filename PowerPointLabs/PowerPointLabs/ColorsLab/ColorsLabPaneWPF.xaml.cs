using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;


using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.DataSources;
using PowerPointLabs.Utils;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

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
            DataContext = dataSource;

            InitializeComponent();

            SetupImageSources();

            SetDefaultColor(Color.CornflowerBlue);
        }


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


        private void SetDefaultColor(Color color)
        {
            dataSource.SelectedColor = color;
            // UpdateUIForNewColor();
        }

        private void BrightnessSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            double newBrightness = e.NewValue;
            HSLColor newColor = new HSLColor();
            newColor.Hue = dataSource.SelectedColor.Hue;
            newColor.Saturation = dataSource.SelectedColor.Saturation;
            newColor.Luminosity = newBrightness;
            dataSource.SelectedColor = newColor;
        }

        private void SaturationSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            double newSaturation = e.NewValue;
            HSLColor newColor = new HSLColor();
            newColor.Hue = dataSource.SelectedColor.Hue;
            newColor.Saturation = newSaturation;
            newColor.Luminosity = dataSource.SelectedColor.Luminosity;
            dataSource.SelectedColor = newColor;
        }
    }
}
