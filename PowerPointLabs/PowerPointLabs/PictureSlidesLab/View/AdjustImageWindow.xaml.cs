using System;
using System.Drawing;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.PictureSlidesLab.Util;
using PowerPointLabs.PictureSlidesLab.View.ImageAdjustment;
using PowerPointLabs.WPF.Observable;
using Color = System.Windows.Media.Color;
using System.Windows.Threading;

namespace PowerPointLabs.PictureSlidesLab.View
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>

    public partial class AdjustImageWindow
    {
        //used for display in the adjust dimensions window only.
        private ObservableString thumbnailImageFile = new ObservableString();
        private const int AdjustUnit = 10;

        private CroppingAdorner _croppingAdorner;
        private FrameworkElement _frameworkElement;

        private bool _isRectSizeSet;
        private double _rectX;
        private double _rectY;
        private double _rectWidth;
        private double _rectHeight;

        //used for applying styles to the ppt
        public string CropResult { get; set; }
        //used in generating previews.
        public string CropResultThumbnail { get; set; }
        public string RotateResult { get; set; }
        public Rect Rect { get; set; }
        public bool IsCropped { get; set; }
        public bool IsRotated { get; set; }
        public bool IsOpen { get; set; }

        public AdjustImageWindow()
        {
            InitializeComponent();
            ImageHolder.DataContext = thumbnailImageFile;
            MoveLeftImage.Source = ImageUtil.BitmapToImageSource(Properties.Resources.Left);
            MoveUpImage.Source = ImageUtil.BitmapToImageSource(Properties.Resources.Up);
            MoveDownImage.Source = ImageUtil.BitmapToImageSource(Properties.Resources.Down);
            MoveRightImage.Source = ImageUtil.BitmapToImageSource(Properties.Resources.Right);
            ZoomInImage.Source = ImageUtil.BitmapToImageSource(Properties.Resources.PlusZoom);
            ZoomOutImage.Source = ImageUtil.BitmapToImageSource(Properties.Resources.MinusZoom);
            AutoFitImage.Source = ImageUtil.BitmapToImageSource(Properties.Resources.Fit);
            LeftRotateImage.Source = ImageUtil.BitmapToImageSource(Properties.Resources.LeftRotate);
            RightRotateImage.Source = ImageUtil.BitmapToImageSource(Properties.Resources.RightRotate);
            FlipHorizontalImage.Source = ImageUtil.BitmapToImageSource(Properties.Resources.FlipHorizontal);
            FlipVerticalImage.Source = ImageUtil.BitmapToImageSource(Properties.Resources.FlipVertical);
        }

        public void ShowAdjustPictureDimensionsDialog()
        {
            IsOpen = true;
            ShowDialog();
            IsOpen = false;
        }

        public void SetThumbnailImage(string imageFile)
        {
            CropResultThumbnail = imageFile;
            Dispatcher.Invoke(new Action(() =>
            {
                thumbnailImageFile.Text = imageFile;
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
                rect = AutoFit(element);
            }

            var layer = AdornerLayer.GetAdornerLayer(element);
            _croppingAdorner = new CroppingAdorner(element, rect);
            _croppingAdorner.SlideWidth = this.GetCurrentPresentation().SlideWidth;
            _croppingAdorner.SlideHeight = this.GetCurrentPresentation().SlideHeight;
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
            // adjust image holder size
            ImageHolder.Width = ImageHolder.ActualWidth;
            ImageHolder.Height = ImageHolder.ActualHeight;
            // adjust this window size
            Height = ImageHolder.ActualHeight + 15 + SaveCropButton.ActualHeight + 50;
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

        private void AutoFitButton_OnClick(object sender, RoutedEventArgs e)
        {
            var rect = AutoFit();
            _croppingAdorner.ClippingRectangle = rect;
        }

        private Rect AutoFit(FrameworkElement element = null)
        {
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            var slideHeight = this.GetCurrentPresentation().SlideHeight;
            element = element ?? _frameworkElement;

            Rect rect;
            if (element.ActualWidth / element.ActualHeight
                < slideWidth/slideHeight)
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
            return rect;
        }

        private void MoveLeftButton_OnClick(object sender, RoutedEventArgs e)
        {
            _croppingAdorner.MoveCroppingRect(-AdjustUnit, 0);
        }

        private void MoveUpButton_OnClick(object sender, RoutedEventArgs e)
        {
            _croppingAdorner.MoveCroppingRect(0, -AdjustUnit);
        }

        private void MoveDownButton_OnClick(object sender, RoutedEventArgs e)
        {
            _croppingAdorner.MoveCroppingRect(0, AdjustUnit);
        }

        private void MoveRightButton_OnClick(object sender, RoutedEventArgs e)
        {
            _croppingAdorner.MoveCroppingRect(AdjustUnit, 0);
        }

        private void ZoomInButton_OnClick(object sender, RoutedEventArgs e)
        {
            _croppingAdorner.ZoomCroppingRect(AdjustUnit);
        }

        private void ZoomOutButton_OnClick(object sender, RoutedEventArgs e)
        {
            _croppingAdorner.ZoomCroppingRect(-AdjustUnit);
        }

        private void RotateFlipImg(RotateFlipType roatateFlipType)
        {
            var img = (Bitmap)Bitmap.FromFile(CropResult);
            img.RotateFlip(roatateFlipType);
            String rotatedImg = StoragePath.GetPath("rotated-"
                                    + DateTime.Now.GetHashCode() + "-"
                                    + Guid.NewGuid().ToString().Substring(0, 7));
            img.Save(rotatedImg);
            CropResult = rotatedImg;
            RotateResult = rotatedImg;
            thumbnailImageFile.Text = ImageUtil.GetThumbnailFromFullSizeImg(rotatedImg);
            CropResultThumbnail = ImageUtil.GetThumbnailFromFullSizeImg(rotatedImg);
            IsRotated = true;

            //Resize the ImageHolder and Window after rotating
            if (roatateFlipType.Equals(RotateFlipType.Rotate270FlipNone)
                || roatateFlipType.Equals(RotateFlipType.Rotate90FlipNone))
            {
                ImageHolder.Width = ImageHolder.ActualHeight / Bitmap.FromFile(CropResultThumbnail).Height * Bitmap.FromFile(CropResultThumbnail).Width;
                ImageHolder.Height = ImageHolder.ActualHeight;
                this.Width = ImageHolder.Width + 20;
                this.CenterWindowOnScreen();
            }

            Dispatcher.Invoke(DispatcherPriority.SystemIdle, new Action(() =>
            {
                var rect = AutoFit();
                _croppingAdorner.ClippingRectangle = rect;
            }));
        }

        private void LeftRotateButton_OnClick(object sender, RoutedEventArgs e)
        {
            RotateFlipImg(RotateFlipType.Rotate270FlipNone);
        }

        private void RightRotateButton_OnClick(object sender, RoutedEventArgs e)
        {
            RotateFlipImg(RotateFlipType.Rotate90FlipNone);
        }

        private void FlipHorizontalButton_OnClick(object sender, RoutedEventArgs e)
        {
            RotateFlipImg(RotateFlipType.Rotate180FlipY);
        }

        private void FlipVerticalButton_OnClick(object sender, RoutedEventArgs e)
        {
            RotateFlipImg(RotateFlipType.Rotate180FlipX);
        }
    }
}