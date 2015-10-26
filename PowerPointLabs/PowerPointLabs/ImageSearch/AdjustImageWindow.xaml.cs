using System;
using System.Drawing;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using DAP.Adorners;
using PowerPointLabs.ImageSearch.Util;
using PowerPointLabs.Models;
using PowerPointLabs.WPF.Observable;
using Color = System.Windows.Media.Color;

namespace PowerPointLabs.ImageSearch
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>

    public partial class AdjustImageWindow
    {
        private ObservableString ThumbnailImageFile = new ObservableString();

        private CroppingAdorner _croppingAdorner;
        private FrameworkElement _frameworkElement;

        private bool _isRectSizeSet;
        private double _rectX;
        private double _rectY;
        private double _rectWidth;
        private double _rectHeight;

        public string CropResult { get; set; }
        public string CropResultThumbnail { get; set; }
        public Rect Rect { get; set; }
        public bool IsCropped { get; set; }
        public bool IsOpen { get; set; }

        public AdjustImageWindow()
        {
            InitializeComponent();
            ImageHolder.DataContext = ThumbnailImageFile;
        }

        public void SetThumbnailImage(string imageFile)
        {
            CropResultThumbnail = imageFile;
            Dispatcher.Invoke(new Action(() =>
            {
                ThumbnailImageFile.Text = imageFile;
            }));
        }

        public void SetFullsizeImage(string imageFile)
        {
            CropResult = imageFile;
        }

        public void SetCropRect(double x, double y, double width, double height)
        {
            _isRectSizeSet = true;
            _rectX = x;
            _rectY = y;
            _rectWidth = width;
            _rectHeight = height;
        }

        private void RemoveCropFromCur()
        {
            AdornerLayer aly = AdornerLayer.GetAdornerLayer(_frameworkElement);
            aly.Remove(_croppingAdorner);
        }

        private void AddCropToElement(FrameworkElement element)
        {
            if (_frameworkElement != null)
            {
                RemoveCropFromCur();
            }

            Rect rect;
            if (_isRectSizeSet)
            {
                rect = new Rect(
                    _rectX,
                    _rectY,
                    _rectWidth,
                    _rectHeight);
            }
            else
            {
                var slideWidth = PowerPointPresentation.Current.SlideWidth;
                var slideHeight = PowerPointPresentation.Current.SlideHeight;
                if (element.ActualWidth / element.ActualHeight
                    < slideWidth / slideHeight)
                {
                    var targetHeight = element.ActualWidth / slideWidth * slideHeight;
                    rect = new Rect(
                        0,
                        (element.ActualHeight - targetHeight) / 2,
                        element.ActualWidth,
                        targetHeight);
                }
                else
                {
                    var targetWidth = element.ActualHeight / slideHeight * slideWidth;
                    rect = new Rect(
                        (element.ActualWidth - targetWidth) / 2,
                        0,
                        targetWidth,
                        element.ActualHeight);
                }
            }

            var layer = AdornerLayer.GetAdornerLayer(element);
            _croppingAdorner = new CroppingAdorner(element, rect);
            _croppingAdorner.SlideWidth = PowerPointPresentation.Current.SlideWidth;
            _croppingAdorner.SlideHeight = PowerPointPresentation.Current.SlideHeight;
            _croppingAdorner.CropChanged += (sender, args) =>
            {
                var croppingRect = _croppingAdorner.ClippingRectangle;
                if (croppingRect.Width*croppingRect.Height < 1)
                {
                    SaveCropButton.IsEnabled = false;
                }
                else
                {
                    SaveCropButton.IsEnabled = true;
                }
            };
            Rect = rect;

            layer.Add(_croppingAdorner);
            _frameworkElement = element;
            SetClipColorGrey();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            AddCropToElement(ImageHolder);
            AdjustControlsSize();
            CenterWindowOnScreen();
        }

        private void AdjustControlsSize()
        {
            ImageHolder.Width = ImageHolder.ActualWidth;
            ImageHolder.Height = ImageHolder.ActualHeight;
            Height = ImageHolder.ActualHeight + 15 + SaveCropButton.ActualHeight + 40;
            Width = ImageHolder.ActualWidth + 20;
        }

        private void CenterWindowOnScreen()
        {
            var screenWidth = SystemParameters.PrimaryScreenWidth;
            var screenHeight = SystemParameters.PrimaryScreenHeight;
            var windowWidth = Width;
            var windowHeight = Height;
            Left = (screenWidth / 2) - (windowWidth / 2);
            Top = (screenHeight / 2) - (windowHeight / 2);
        }

        private void SetClipColorGrey()
        {
            if (_croppingAdorner != null)
            {
                Color color = Colors.Black;
                color.A = 110;
                _croppingAdorner.Fill = new SolidColorBrush(color);
            }
        }

        private void SaveCropButton_OnClick(object sender, RoutedEventArgs e)
        {
            var rect = _croppingAdorner.ClippingRectangle;
            var xRatio = rect.X/ImageHolder.ActualWidth;
            var yRatio = rect.Y/ImageHolder.ActualHeight;
            var widthRatio = rect.Width/ImageHolder.ActualWidth;
            var heightRatio = rect.Height/ImageHolder.ActualHeight;
            Rect = rect;

            var originalImg = (Bitmap)Bitmap.FromFile(CropResult);
            var result = CropToShape.KiCut(originalImg, 
                (float) xRatio*originalImg.Width,
                (float) yRatio * originalImg.Height,
                (float) widthRatio * originalImg.Width,
                (float) heightRatio * originalImg.Height);
            CropResult = StoragePath.GetPath("crop-"
                                    + DateTime.Now.GetHashCode() + "-"
                                    + Guid.NewGuid().ToString().Substring(0, 7));
            result.Save(CropResult);

            CropResultThumbnail = ImageUtil.GetThumbnailFromFullSizeImg(CropResult);
            IsCropped = true;

            Close();
        }
    }
}