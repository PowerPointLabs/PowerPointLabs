using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
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
        public static float PictureExportingRatio = 330.0f / 72.0f;
        private const float targetDpi = 96.0f;
        private static float dpiScale = 1.0f;

        // Picture exporting ratios
        private const float pictureExportingRatioHigh = 330.0f / 72.0f;
        private const float pictureExportingRatioCompressed = 96.0f / 72.0f;

        // Heuristics for image compression obtained through testing
        private const long targetCompression = 75L;
        private const long fileSizeLimit = 40000L;

        // Static initializer to retrieve dpi scale once
        static GraphicsUtil()
        {
            using (Graphics g = Graphics.FromHwnd(IntPtr.Zero))
            {
                dpiScale = g.DpiX / targetDpi;
            }
        }
        #endregion

        #region API

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
            int slideWidth = (int)PowerPointPresentation.Current.SlideWidth;
            int slideHeight = (int)PowerPointPresentation.Current.SlideHeight;

            shapeRange.Export(exportPath, PpShapeFormat.ppShapeFormatPNG, slideWidth,
                              slideHeight, PpExportMode.ppScaleToFit);
        }

        public static Shape CutAndPaste(Shape shape, Slide slide)
        {
            string tempFilePath = FileDir.GetTemporaryPngFilePath();
            ExportShape(shape, tempFilePath);
            Shape resultShape = ImportPictureToSlide(shape, slide, tempFilePath);
            FileDir.TryDeleteFile(tempFilePath);
            return resultShape;
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

        public static Shape CompressImageInShape(Shape targetShape, PowerPointSlide currentSlide)
        {
            // Specify the temp location to be saved. Must be cleared before each new access to that location
            string fileName = CommonText.TemporaryCompressedImageStorageFileName;
            string tempFileStoragePath = Path.Combine(Path.GetTempPath(), fileName);

            // Export the shape to extract the image within it
            targetShape.Export(tempFileStoragePath, PpShapeFormat.ppShapeFormatJPG);

            // Check if the image is acceptable in terms of size
            FileInfo tempFile = new FileInfo(tempFileStoragePath);
            tempFile.Refresh();
            long fileLength = tempFile.Length;
            if (fileLength < fileSizeLimit)
            {
                // Delete the file as it is not needed anymore
                DeleteSpecificFilePath(tempFileStoragePath);

                // Return the original shape
                return targetShape;
            }

            // Create a new bitmap from the image representing the exported shape
            Image img = Image.FromFile(tempFileStoragePath);
            Bitmap imgBitMap = new Bitmap(img);

            // Releases resources held by image object and delete temp file
            img.Dispose();
            DeleteSpecificFilePath(tempFileStoragePath);

            // Compresses and save the bitmap based on the specified level of compression (0 -> lowest quality, 100 -> highest quality)
            SaveJpg(imgBitMap, tempFileStoragePath, targetCompression);

            // Retrieve the compressed image and return a shape representing the image
            Shape compressedImgShape = currentSlide.Shapes.AddPicture(tempFileStoragePath,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    targetShape.Left,
                    targetShape.Top);

            // Delete temp file again to return to original empty state
            DeleteSpecificFilePath(tempFileStoragePath);

            // Delete targetShape to prevent duplication
            targetShape.Delete();

            return compressedImgShape;
        }

        // Save the file with a specific compression level.
        private static void SaveJpg(Bitmap bm, string fileName, long compression)
        {
            System.Drawing.Imaging.EncoderParameters encoderParams = new System.Drawing.Imaging.EncoderParameters(1);
            encoderParams.Param[0] = new System.Drawing.Imaging.EncoderParameter(
                System.Drawing.Imaging.Encoder.Quality, compression);

            System.Drawing.Imaging.ImageCodecInfo imageCodecInfo = GetEncoderInfo("image/jpeg");
            File.Delete(fileName);
            bm.Save(fileName, imageCodecInfo, encoderParams);
        }

        private static System.Drawing.Imaging.ImageCodecInfo GetEncoderInfo(string mimeType)
        {
            System.Drawing.Imaging.ImageCodecInfo[] encoders = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders();
            for (int i = 0; i <= encoders.Length; i++)
            {
                if (encoders[i].MimeType == mimeType)
                {
                    return encoders[i];
                }
            }
            return null;
        }

        private static void DeleteSpecificFilePath(string filePath)
        {
            FileInfo file = new FileInfo(filePath);
            if (file.Exists)
            {
                file.Delete();
            }
        }

        #endregion

        #region Slide
        public static void ExportSlides(List<Slide> slides, string exportPath, float magnifyRatio = 1.0f)
        {
            if (slides.Count <= 0)
            {
                return;
            } 
            else if (slides.Count == 1)
            {
                ExportSlide(slides[0], exportPath, magnifyRatio);
                return;
            }

            // Get folder name from exportPath
            string folderName = GetDefaultFolderNameForExport(exportPath);

            try
            {
                Directory.CreateDirectory(folderName);

                foreach (Slide slide in slides)
                {
                    string fileName = folderName + "\\" + slide.Name + ".png";
                    ExportSlide(slide, fileName, magnifyRatio);
                }

                // Alert the user that the slides have been saved in a folder
                string messageBoxText = "The selected slides have been saved as a separate file in the folder " + folderName + ".";
                MessageBox.Show(messageBoxText);
            }
            catch (Exception)
            {
                // Failed to create directory, we save the images all to the specified path
                foreach (Slide slide in slides)
                {
                    string fileName = folderName + "_" + slide.Name + ".png";
                    ExportSlide(slide, fileName, magnifyRatio);
                }

                // Alert the user that the slides have been saved as a separate file in the specified folder.
                string messageBoxText = "The selected slides have been saved as a separate file in the specified folder.";
                MessageBox.Show(messageBoxText);
            }
        }

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

        public static Shape AddSlideAsShape(PowerPointSlide slideToAdd, PowerPointSlide targetSlide)
        {
            string tempFilePath = FileDir.GetTemporaryPngFilePath();
            ExportSlide(slideToAdd, tempFilePath);
            Shape slideAsShape = ImportPictureToSlide(slideToAdd, targetSlide, tempFilePath);
            FileDir.TryDeleteFile(tempFilePath);
            return slideAsShape;
        }

        public static Shape AddAudioShapeFromFile(PowerPointSlide slide, string fileName)
        {
            float slideWidth = PowerPointPresentation.Current.SlideWidth;
            Shape audioShape = slide.Shapes.AddMediaObject2(fileName, MsoTriState.msoFalse, MsoTriState.msoTrue, slideWidth + 20);
            if (audioShape == null)
            {
                MessageBox.Show("The audio file cannot be accessed by PowerPoint. " +
                    "Try moving it to another location like Documents or Desktop.");
                return null;
            }
            try
            {
                slide.RemoveAnimationsForShape(audioShape);
                return audioShape;
            }
            catch
            {
                Logger.Log("Audio not generated because text is not in English.");
                return null;
            }
        }

        private static Shape ImportPictureToSlide(Shape shapeToAdd, Slide targetSlide, string tempFilePath)
        {
            // The AccessViolationException is longer catchable
            if (!FileDir.IsFileReadable(tempFilePath))
            {
                shapeToAdd.Cut();
                return targetSlide.Shapes.PasteSpecial(PpPasteDataType.ppPastePNG)[1];
            }
            else
            {
                shapeToAdd.Delete();
                return targetSlide.Shapes.AddPicture2(tempFilePath,
                    MsoTriState.msoFalse,
                    MsoTriState.msoTrue,
                    0,
                    0);
            }
        }

        private static Shape ImportPictureToSlide(PowerPointSlide slideToAdd, PowerPointSlide targetSlide, string tempFilePath)
        {
            // The AccessViolationException is longer catchable
            if (!FileDir.IsFileReadable(tempFilePath))
            {
                return targetSlide.Shapes.SafeCopySlide(slideToAdd);
            }
            else
            {
                return targetSlide.Shapes.AddPicture2(tempFilePath,
                                                      MsoTriState.msoFalse,
                                                      MsoTriState.msoTrue,
                                                      0,
                                                      0);
            }
        }

        #endregion

        #region Bitmap
        public static Bitmap CreateThumbnailImage(Image oriImage, int width, int height)
        {
            double scalingRatio = CalculateScalingRatio(oriImage.Size, new Size(width, height));

            // calculate width and height after scaling
            int scaledWidth = (int)Math.Round(oriImage.Size.Width * scalingRatio);
            int scaledHeight = (int)Math.Round(oriImage.Size.Height * scalingRatio);

            // calculate left top corner position of the image in the thumbnail
            int scaledLeft = (width - scaledWidth) / 2;
            int scaledTop = (height - scaledHeight) / 2;

            // define drawing area
            Rectangle drawingRect = new Rectangle(scaledLeft, scaledTop, scaledWidth, scaledHeight);
            Bitmap thumbnail = new Bitmap(width, height);

            // here we set the thumbnail as the highest quality
            using (Graphics thumbnailGraphics = System.Drawing.Graphics.FromImage(thumbnail))
            {
                thumbnailGraphics.CompositingQuality = CompositingQuality.HighQuality;
                thumbnailGraphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                thumbnailGraphics.SmoothingMode = SmoothingMode.HighQuality;
                thumbnailGraphics.DrawImage(oriImage, drawingRect);
            }

            return thumbnail;
        }
        #endregion

        #region GDI+
        public static void SuspendDrawing(Control control)
        {
            Native.SendMessage(control.Handle, (uint) Native.Message.WM_SETREDRAW, IntPtr.Zero, IntPtr.Zero);
        }

        public static void ResumeDrawing(Control control)
        {
            Native.SendMessage(control.Handle, (uint) Native.Message.WM_SETREDRAW, new IntPtr(1), IntPtr.Zero);
            control.Refresh();
        }
        #endregion
        
        #region Color
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

        #region Settings

        public static void ShouldCompressPictureExport(bool shouldCompress)
        {
            PictureExportingRatio = shouldCompress ? pictureExportingRatioCompressed : pictureExportingRatioHigh;
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
            // Powerpoint displays at 72 dpi, while the picture stores in 96 dpi or 330 dpi, depending on user option.
            return PowerPointPresentation.Current.SlideWidth * PictureExportingRatio;
        }

        private static double GetDesiredExportHeight()
        {
            // Powerpoint displays at 72 dpi, while the picture stores in 96 dpi or 330 dpi, depending on user option.
            return PowerPointPresentation.Current.SlideHeight * PictureExportingRatio;
        }

        private static string GetDefaultFolderNameForExport(string exportPath)
        {
            string folderName = exportPath.Substring(0, exportPath.Length - 4);

            int suffix = 1;
            int idx = exportPath.LastIndexOf('\\');
            string folderPath = exportPath.Substring(0, idx + 1);

            while (Directory.Exists(folderName))
            {
                // Change to default folder name with suffix
                string suffixString = suffix.ToString();
                folderName = folderPath + "PPTLabs_ExportedSlides_" + suffixString;
                suffix++;
            }

            return folderName;
        }

        /// <summary>
        /// Converts a Bitmap to Bitmap source
        /// </summary>
        /// <param name="bitmap">The bitmap to convert</param>
        /// <returns>The converted object</returns>
        public static BitmapSource CreateBitmapSourceFromGdiBitmap(Bitmap bitmap)
        {
            Rectangle rect = new System.Drawing.Rectangle(0, 0, bitmap.Width, bitmap.Height);

            BitmapData bitmapData = bitmap.LockBits(
                rect,
                ImageLockMode.ReadWrite,
                Drawing.Imaging.PixelFormat.Format32bppArgb);

            try
            {
                int size = (rect.Width * rect.Height) * 4;

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
        #endregion
    }
}
