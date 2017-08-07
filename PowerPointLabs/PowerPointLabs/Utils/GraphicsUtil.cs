using System;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Imaging;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;

using PPExtraEventHelper;

using Drawing = System.Drawing;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;

namespace PowerPointLabs.Utils
{
    [SuppressMessage("Microsoft.StyleCop.CSharp.OrderingRules", "SA1202:ElementsMustBeOrderedByAccess", Justification = "To refactor to partials")]
    internal static class GraphicsUtil
    {
#pragma warning disable 0618

        #region Constants
        private static readonly Object FileLock = new object();
        public const float PictureExportingRatio = 96.0f / 72.0f;
        private const float TargetDpi = 96.0f;
        private static float dpiScale = 1.0f;

        // Static initializer to retrieve dpi scale once
        static GraphicsUtil()
        {
            using (Graphics g = Graphics.FromHwnd(IntPtr.Zero))
            {
                dpiScale = g.DpiX / TargetDpi;
            }
        }
        # endregion

        # region API

        # region Clipboard
        
        public static bool IsClipboardEmpty()
        {
            IDataObject clipboardData = Clipboard.GetDataObject();
            return clipboardData == null || clipboardData.GetFormats().Length == 0;
        }

        #endregion

        #region Shape

        public static void ExportShape(Shape shape, string exportPath)
        {
            int slideWidth = 0;
            int slideHeight = 0;
            try
            {
                slideWidth = (int)PowerPointPresentation.Current.SlideWidth;
                slideHeight = (int)PowerPointPresentation.Current.SlideHeight;
            }
            catch (NullReferenceException)
            {
                // Getting Presentation.Current may throw NullReferenceException during unit testing
                shape.Export(exportPath, PpShapeFormat.ppShapeFormatPNG, ExportMode: PpExportMode.ppScaleToFit);
            }

            shape.Export(exportPath, PpShapeFormat.ppShapeFormatPNG, slideWidth,
                         slideHeight, PpExportMode.ppScaleToFit);
        }

        public static void ExportShape(ShapeRange shapeRange, string exportPath)
        {
            var slideWidth = (int)PowerPointPresentation.Current.SlideWidth;
            var slideHeight = (int)PowerPointPresentation.Current.SlideHeight;

            shapeRange.Export(exportPath, PpShapeFormat.ppShapeFormatPNG, slideWidth,
                              slideHeight, PpExportMode.ppScaleToFit);
        }

        public static Bitmap ShapeToBitmap(Shape shape)
        {
            // we need a lock here to prevent race conditions on the temporary file
            lock (FileLock)
            {
                string fileName = CommonText.TemporaryImageStorageFileName;
                string tempPicPath = Path.Combine(Path.GetTempPath(), fileName);
                ExportShape(shape, tempPicPath);

                Image image = Image.FromFile(tempPicPath);
                Bitmap bitmap = new Bitmap(image);
                // free up the original file to be deleted
                image.Dispose();

                FileInfo file = new FileInfo(Path.GetTempPath() + fileName);
                if (file.Exists)
                {
                    file.Delete();
                }
                return bitmap;
            }
        }

        #endregion

        # region Slide
        public static void ExportSlide(Slide slide, string exportPath, float magnifyRatio = 1.0f)
        {
            slide.Export(exportPath,
                         "PNG",
                         (int)(GetDesiredExportWidth() * magnifyRatio),
                         (int)(GetDesiredExportHeight() * magnifyRatio));
        }

        public static void ExportSlide(PowerPointSlide slide, string exportPath, float magnifyRatio = 1.0f)
        {
            ExportSlide(slide.GetNativeSlide(), exportPath, magnifyRatio);
        }

        # endregion

        # region Bitmap
        public static Bitmap CreateThumbnailImage(Image oriImage, int width, int height)
        {
            var scalingRatio = CalculateScalingRatio(oriImage.Size, new Size(width, height));

            // calculate width and height after scaling
            var scaledWidth = (int)Math.Round(oriImage.Size.Width * scalingRatio);
            var scaledHeight = (int)Math.Round(oriImage.Size.Height * scalingRatio);

            // calculate left top corner position of the image in the thumbnail
            var scaledLeft = (width - scaledWidth) / 2;
            var scaledTop = (height - scaledHeight) / 2;

            // define drawing area
            var drawingRect = new Rectangle(scaledLeft, scaledTop, scaledWidth, scaledHeight);
            var thumbnail = new Bitmap(width, height);

            // here we set the thumbnail as the highest quality
            using (var thumbnailGraphics = System.Drawing.Graphics.FromImage(thumbnail))
            {
                thumbnailGraphics.CompositingQuality = CompositingQuality.HighQuality;
                thumbnailGraphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                thumbnailGraphics.SmoothingMode = SmoothingMode.HighQuality;
                thumbnailGraphics.DrawImage(oriImage, drawingRect);
            }

            return thumbnail;
        }
        # endregion

        # region GDI+
        public static void SuspendDrawing(Control control)
        {
            Native.SendMessage(control.Handle, (uint) Native.Message.WM_SETREDRAW, IntPtr.Zero, IntPtr.Zero);
        }

        public static void ResumeDrawing(Control control)
        {
            Native.SendMessage(control.Handle, (uint) Native.Message.WM_SETREDRAW, new IntPtr(1), IntPtr.Zero);
            control.Refresh();
        }
        # endregion
        
        # region Color
        public static int ConvertColorToRgb(Drawing.Color argb)
        {
            return (argb.B << 16) | (argb.G << 8) | argb.R;
        }

        public static int PackRgbInt(byte r, int g, int b)
        {
            return (b << 16) | (g << 8) | r;
        }

        public static Drawing.Color ConvertRgbToColor(int rgb)
        {
            return Drawing.Color.FromArgb(rgb & 255, (rgb >> 8) & 255, (rgb >> 16) & 255);
        }

        public static void UnpackRgbInt(int rgb, out byte r, out byte g, out byte b)
        {
            r = (byte)(rgb & 255);
            g = (byte)((rgb >> 8) & 255);
            b = (byte)((rgb >> 16) & 255);
        }

        public static Drawing.Color DrawingColorFromMediaColor(System.Windows.Media.Color mediaColor)
        {
            return Drawing.Color.FromArgb(mediaColor.A, mediaColor.R, mediaColor.G, mediaColor.B);
        }

        public static System.Windows.Media.Color MediaColorFromDrawingColor(Drawing.Color drawingColor)
        {
            return System.Windows.Media.Color.FromArgb(drawingColor.A, drawingColor.R, drawingColor.G, drawingColor.B);
        }

        public static Drawing.Color DrawingColorFromBrush(System.Windows.Media.Brush brush)
        {
            return DrawingColorFromMediaColor((brush as SolidColorBrush).Color);
        }

        public static System.Windows.Media.Brush MediaBrushFromDrawingColor(Drawing.Color color)
        {
            return new SolidColorBrush(MediaColorFromDrawingColor(color));
        }
        #endregion

        #endregion

        #region Helper Methods
        private static double CalculateScalingRatio(Size oldSize, Size newSize)
        {
            double scalingRatio;

            if (oldSize.Width >= oldSize.Height)
            {
                scalingRatio = (double)newSize.Width / oldSize.Width;
            }
            else
            {
                scalingRatio = (double)newSize.Height / oldSize.Height;
            }

            return scalingRatio;
        }

        private static double GetDesiredExportWidth()
        {
            // Powerpoint displays at 72 dpi, while the picture stores in 96 dpi.
            return PowerPointPresentation.Current.SlideWidth / 72.0 * 96.0;
        }

        private static double GetDesiredExportHeight()
        {
            // Powerpoint displays at 72 dpi, while the picture stores in 96 dpi.
            return PowerPointPresentation.Current.SlideHeight / 72.0 * 96.0;
        }

        /// <summary>
        /// Converts a Bitmap to Bitmap source
        /// </summary>
        /// <param name="bitmap">The bitmap to convert</param>
        /// <returns>The converted object</returns>
        public static BitmapSource CreateBitmapSourceFromGdiBitmap(Bitmap bitmap)
        {
            var rect = new System.Drawing.Rectangle(0, 0, bitmap.Width, bitmap.Height);

            var bitmapData = bitmap.LockBits(
                rect,
                ImageLockMode.ReadWrite,
                Drawing.Imaging.PixelFormat.Format32bppArgb);

            try
            {
                var size = (rect.Width * rect.Height) * 4;

                return BitmapSource.Create(
                    bitmap.Width,
                    bitmap.Height,
                    bitmap.HorizontalResolution,
                    bitmap.VerticalResolution,
                    PixelFormats.Bgra32,
                    null,
                    bitmapData.Scan0,
                    size,
                    bitmapData.Stride);
            }
            finally
            {
                bitmap.UnlockBits(bitmapData);
            }
        }

        public static float GetDpiScale()
        {
            return dpiScale;
        }
        # endregion
    }
}
