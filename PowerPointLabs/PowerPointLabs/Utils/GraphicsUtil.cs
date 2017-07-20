using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Imaging;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;

using PPExtraEventHelper;

using Drawing = System.Drawing;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;
using TextFrame2 = Microsoft.Office.Interop.PowerPoint.TextFrame2;

namespace PowerPointLabs.Utils
{
    [SuppressMessage("Microsoft.StyleCop.CSharp.OrderingRules", "SA1202:ElementsMustBeOrderedByAccess", Justification = "To refactor to partials")]
    internal static class GraphicsUtil
    {
#pragma warning disable 0618

        #region Const
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
                string fileName = TextCollection.TemporaryImageStorageFileName;
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

        #region Text
        public static TextRange ConvertTextRange2ToTextRange(TextRange2 textRange2)
        {
            var textFrame2 = textRange2.Parent as TextFrame2;

            if (textFrame2 == null)
            {
                return null;
            }

            var shape = textFrame2.Parent as Shape;

            return shape == null ? null : shape.TextFrame.TextRange;
        }
        # endregion

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

        /// <summary>
        /// Sort by increasing index.
        /// </summary>
        public static void SortByIndex(List<PowerPointSlide> slides)
        {
            slides.Sort((sh1, sh2) => sh1.Index - sh2.Index);
        }

        /// <summary>
        /// Used for the SquashSlides method.
        /// This struct holds transition information for an effect.
        /// </summary>
        private struct EffectTransition
        {
            private readonly MsoAnimTriggerType slideTransition;
            private readonly float transitionTime;

            public EffectTransition(MsoAnimTriggerType slideTransition, float transitionTime)
            {
                this.slideTransition = slideTransition;
                this.transitionTime = transitionTime;
            }

            public void ApplyTransition(Effect effect)
            {
                effect.Timing.TriggerType = slideTransition;
                effect.Timing.TriggerDelayTime = transitionTime;
            }
        }

        /// <summary>
        /// Merges multiple animated slides into a single slide.
        /// TODO: Test this method more thoroughly, in places other than autozoom.
        /// </summary>
        public static void SquashSlides(IEnumerable<PowerPointSlide> slides)
        {
            PowerPointSlide firstSlide = null;
            List<Shape> previousShapes = null;
            EffectTransition slideTransition = new EffectTransition();

            foreach (var slide in slides)
            {
                if (firstSlide == null)
                {
                    firstSlide = slide;
                    slideTransition = GetTransitionFromSlide(slide);

                    firstSlide.Transition.AdvanceOnClick = MsoTriState.msoTrue;
                    firstSlide.Transition.AdvanceOnTime = MsoTriState.msoFalse;

                    //TODO: Check if there will be an exception when there are empty placeholder shapes in firstSlide.
                    previousShapes = firstSlide.Shapes.Cast<Shape>().ToList();
                    continue;
                }

                var effectSequence = firstSlide.GetNativeSlide().TimeLine.MainSequence;
                int effectStartIndex = effectSequence.Count + 1;


                slide.DeleteIndicator();
                var newShapeRange = firstSlide.CopyShapesToSlide(slide.Shapes.Range());
                newShapeRange.ZOrder(MsoZOrderCmd.msoSendToBack);


                var newShapes = newShapeRange.Cast<Shape>().ToList();
                newShapes.ForEach(shape => AddAppearAnimation(shape, firstSlide, effectStartIndex));
                previousShapes.ForEach(shape => AddDisappearAnimation(shape, firstSlide, effectStartIndex));
                slideTransition.ApplyTransition(effectSequence[effectStartIndex]);


                previousShapes = newShapes;
                slideTransition = GetTransitionFromSlide(slide);
                slide.Delete();
            }
        }

        /// <summary>
        /// Extracts the transition animation out of slide to be used as a transition animation for shapes.
        /// For now, it only extracts the trigger type (trigger by wait or by mouse click), not actual slide transitions.
        /// </summary>
        private static EffectTransition GetTransitionFromSlide(PowerPointSlide slide)
        {
            var transition = slide.GetNativeSlide().SlideShowTransition;
            
            if (transition.AdvanceOnTime == MsoTriState.msoTrue)
            {
                return new EffectTransition(MsoAnimTriggerType.msoAnimTriggerAfterPrevious, transition.AdvanceTime);
            }
            return new EffectTransition(MsoAnimTriggerType.msoAnimTriggerOnPageClick, 0);
        }

        private static void AddDisappearAnimation(Shape shape, PowerPointSlide inSlide, int effectStartIndex)
        {
            if (inSlide.HasExitAnimation(shape))
            {
                return;
            }

            var effectFade = inSlide.GetNativeSlide().TimeLine.MainSequence.AddEffect(shape, MsoAnimEffect.msoAnimEffectAppear,
                MsoAnimateByLevel.msoAnimateLevelNone, MsoAnimTriggerType.msoAnimTriggerWithPrevious, effectStartIndex);
            effectFade.Exit = MsoTriState.msoTrue;
        }

        private static void AddAppearAnimation(Shape shape, PowerPointSlide inSlide, int effectStartIndex)
        {
            if (inSlide.HasEntryAnimation(shape))
            {
                return;
            }

            var effectFade = inSlide.GetNativeSlide().TimeLine.MainSequence.AddEffect(shape, MsoAnimEffect.msoAnimEffectAppear,
                MsoAnimateByLevel.msoAnimateLevelNone, MsoAnimTriggerType.msoAnimTriggerWithPrevious, effectStartIndex);
            effectFade.Exit = MsoTriState.msoFalse;
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

        # region Design

        public static Design CreateDesign(string designName)
        {
            return PowerPointPresentation.Current.Presentation.Designs.Add(designName);
        }

        public static Design GetDesign(string designName)
        {
            foreach (Design design in PowerPointPresentation.Current.Presentation.Designs)
            {
                if (design.Name.Equals(designName))
                {
                    return design;
                }
            }
            return null;
        }

        public static void CopyToDesign(string designName, PowerPointSlide refSlide)
        {
            var design = GetDesign(designName);
            if (design == null)
            {
                design = CreateDesign(designName);
            }
            design.SlideMaster.Background.Fill.ForeColor = refSlide.GetNativeSlide().Background.Fill.ForeColor;
            design.SlideMaster.Background.Fill.BackColor = refSlide.GetNativeSlide().Background.Fill.BackColor;
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

        #region Transformations
        public static PointF RotatePoint(PointF p, PointF origin, float rotation)
        {
            var rotationInRadian = DegreeToRadian(rotation);
            var rotatedX = Math.Cos(rotationInRadian) * (p.X - origin.X) - Math.Sin(rotationInRadian) * (p.Y - origin.Y) + origin.X;
            var rotatedY = Math.Sin(rotationInRadian) * (p.X - origin.X) + Math.Cos(rotationInRadian) * (p.Y - origin.Y) + origin.Y;

            return new PointF((float)rotatedX, (float)rotatedY);
        }

        public static double DegreeToRadian(float angle)
        {
            return angle / 180.0 * Math.PI;
        }
        #endregion

        #endregion

        #region Helper Functions
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
