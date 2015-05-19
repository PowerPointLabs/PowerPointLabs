using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PPExtraEventHelper;
using PowerPointLabs.Models;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;
using TextFrame2 = Microsoft.Office.Interop.PowerPoint.TextFrame2;

namespace PowerPointLabs.Utils
{
    public static class Graphics
    {
        # region Const
        public const float PictureExportingRatio = 96.0f / 72.0f;
        # endregion

        # region API
        # region Shape
        public static Shape CorruptionCorrection(Shape shape, PowerPointSlide ownerSlide)
        {
            // in case of random corruption of shape, cut-paste a shape before using its property
            shape.Cut();
            return ownerSlide.Shapes.Paste()[1];
        }

        public static void ExportShape(Shape shape, string exportPath)
        {
            var slideWidth = (int)PowerPointPresentation.Current.SlideWidth;
            var slideHeight = (int)PowerPointPresentation.Current.SlideHeight;

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

        public static void FitShapeToSlide(ref Shape shapeToMove)
        {
            shapeToMove.LockAspectRatio = MsoTriState.msoFalse;
            shapeToMove.Left = 0;
            shapeToMove.Top = 0;
            shapeToMove.Width = PowerPointPresentation.Current.SlideWidth;
            shapeToMove.Height = PowerPointPresentation.Current.SlideHeight;
        }

        public static bool IsCorrupted(Shape shape)
        {
            try
            {
                return shape.Parent == null;
            }
            catch (Exception)
            {
                return true;
            }
        }

        public static bool IsSamePosition(Shape refShape, Shape candidateShape,
                                          bool exactMatch = true, float blurRadius = float.Epsilon)
        {
            if (exactMatch)
            {
                blurRadius = float.Epsilon;
            }

            return refShape != null &&
                   candidateShape != null &&
                   Math.Abs(refShape.Left - candidateShape.Left) < blurRadius &&
                   Math.Abs(refShape.Top - candidateShape.Top) < blurRadius;
        }

        public static bool IsSameSize(Shape refShape, Shape candidateShape,
                                      bool exactMatch = true, float blurRadius = float.Epsilon)
        {
            if (exactMatch)
            {
                blurRadius = float.Epsilon;
            }

            return refShape != null &&
                   candidateShape != null && 
                   Math.Abs(refShape.Width - candidateShape.Width) < blurRadius &&
                   Math.Abs(refShape.Height - candidateShape.Height) < blurRadius;
        }

        public static bool IsSameType(Shape refShape, Shape candidateShape)
        {
            return refShape != null &&
                   candidateShape != null && 
                   refShape.Type == candidateShape.Type &&
                   (refShape.Type != MsoShapeType.msoAutoShape ||
                   refShape.AutoShapeType == candidateShape.AutoShapeType);
        }

        public static void MakeShapeViewTimeInvisible(Shape shape, Slide curSlide)
        {
            var sequence = curSlide.TimeLine.MainSequence;

            var effectAppear = sequence.AddEffect(shape, MsoAnimEffect.msoAnimEffectAppear,
                                                  MsoAnimateByLevel.msoAnimateLevelNone,
                                                  MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectAppear.Timing.Duration = 0;

            var effectDisappear = sequence.AddEffect(shape, MsoAnimEffect.msoAnimEffectAppear,
                                                     MsoAnimateByLevel.msoAnimateLevelNone,
                                                     MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Exit = MsoTriState.msoTrue;
            effectDisappear.Timing.Duration = 0;

            effectAppear.MoveTo(1);
            effectDisappear.MoveTo(2);
        }

        public static void MakeShapeViewTimeInvisible(Shape shape, PowerPointSlide curSlide)
        {
            MakeShapeViewTimeInvisible(shape, curSlide.GetNativeSlide());
        }

        public static void MakeShapeViewTimeInvisible(ShapeRange shapeRange, Slide curSlide)
        {
            foreach (Shape shape in shapeRange)
            {
                MakeShapeViewTimeInvisible(shape, curSlide);
            }
        }

        public static void MakeShapeViewTimeInvisible(ShapeRange shapeRange, PowerPointSlide curSlide)
        {
            MakeShapeViewTimeInvisible(shapeRange, curSlide.GetNativeSlide());
        }

        public static void SyncShape(Shape refShape, Shape candidateShape,
                                     bool pickupShapeBasic = true, bool pickupShapeFormat = true,
                                     bool pickupTextContent = true, bool pickupTextFormat = true)
        {
            if (pickupShapeBasic)
            {
                SyncShapeRotation(refShape, candidateShape);
                SyncShapeSize(refShape, candidateShape);
                SyncShapeLocation(refShape, candidateShape);
            }

            if ((pickupTextContent || pickupTextFormat) &&
                refShape.HasTextFrame == MsoTriState.msoTrue &&
                candidateShape.HasTextFrame == MsoTriState.msoTrue)
            {
                var refParagraphCount = refShape.TextFrame2.TextRange.Paragraphs.Count;
                var candidateParagraphCount = candidateShape.TextFrame2.TextRange.Paragraphs.Count;
                var refTextRange = refShape.TextFrame2.TextRange;
                var candidateTextRange = candidateShape.TextFrame2.TextRange;

                for (var i = 1; i <= candidateParagraphCount; i++)
                {
                    var refParagraph = refTextRange.Paragraphs[i < refParagraphCount ? i : refParagraphCount];
                    var candidateParagraph = candidateTextRange.Paragraphs[i];

                    SyncTextRange(refParagraph, candidateParagraph, pickupTextContent, pickupTextFormat);
                }
            }

            if (pickupShapeFormat)
            {
                refShape.PickUp();
                candidateShape.Apply();
            }
        }

        public static void SyncShapeRange(ShapeRange refShapeRange, ShapeRange candidateShapeRange)
        {
            // all names of identical shapes should be consistent
            if (refShapeRange.Count != candidateShapeRange.Count)
            {
                return;
            }

            foreach (var shape in candidateShapeRange)
            {
                var candidateShape = shape as Shape;
                var refShape = refShapeRange.Cast<Shape>().FirstOrDefault(item => IsSameType(item, candidateShape) &&
                                                                                  IsSamePosition(item, candidateShape,
                                                                                                 false, 15) &&
                                                                                  IsSameSize(item, candidateShape));

                if (candidateShape == null || refShape == null) continue;

                candidateShape.Name = refShape.Name;
            }
        }

        public static void SyncTextRange(TextRange2 refTextRange, TextRange2 candidateTextRange,
                                         bool pickupTextContent = true, bool pickupTextFormat = true)
        {
            var candidateText = candidateTextRange.Text.TrimEnd('\r');
            var newCandidateRange = candidateTextRange;

            if (pickupTextFormat)
            {
                // pick up format using copy-paste, since we could not deep copy the format
                refTextRange.Copy();
                newCandidateRange = candidateTextRange.PasteSpecial(MsoClipboardFormat.msoClipboardFormatNative);
            }

            if (!pickupTextContent)
            {
                newCandidateRange.Text = candidateText;
            }
        }

        /// <summary>
        /// Moves shiftShape forward/backward until it is just behind destinationShape
        /// </summary>
        public static void MoveZToJustBehind(Shape shiftShape, Shape destinationShape)
        {
            // Step 1: Shift forward until it overshoots destination.
            while (shiftShape.ZOrderPosition < destinationShape.ZOrderPosition)
            {
                int currentValue = shiftShape.ZOrderPosition;
                shiftShape.ZOrder(MsoZOrderCmd.msoBringForward);
                if (shiftShape.ZOrderPosition == currentValue)
                {
                    // Break if no change. Guards against infinite loops.
                    break;
                }
            }

            // Step 2: Shift backward until it overshoots destination.
            while (shiftShape.ZOrderPosition > destinationShape.ZOrderPosition)
            {
                int currentValue = shiftShape.ZOrderPosition;
                shiftShape.ZOrder(MsoZOrderCmd.msoSendBackward);
                if (shiftShape.ZOrderPosition == currentValue)
                {
                    // Break if no change. Guards against infinite loops.
                    break;
                }
            }
        }

        /// <summary>
        /// Moves shiftShape forward/backward until it is just in front of destinationShape
        /// </summary>
        public static void MoveZToJustInFront(Shape shiftShape, Shape destinationShape)
        {
            // Step 1: Shift backward until it overshoots destination.
            while (shiftShape.ZOrderPosition > destinationShape.ZOrderPosition)
            {
                int currentValue = shiftShape.ZOrderPosition;
                shiftShape.ZOrder(MsoZOrderCmd.msoSendBackward);
                if (shiftShape.ZOrderPosition == currentValue)
                {
                    // Break if no change. Guards against infinite loops.
                    break;
                }
            }

            // Step 2: Shift forward until it overshoots destination.
            while (shiftShape.ZOrderPosition < destinationShape.ZOrderPosition)
            {
                int currentValue = shiftShape.ZOrderPosition;
                shiftShape.ZOrder(MsoZOrderCmd.msoBringForward);
                if (shiftShape.ZOrderPosition == currentValue)
                {
                    // Break if no change. Guards against infinite loops.
                    break;
                }
            }
        }

        # endregion

        # region Text
        public static TextRange ConvertTextRange2ToTextRange(TextRange2 textRange2)
        {
            var textFrame2 = textRange2.Parent as TextFrame2;

            if (textFrame2 == null) return null;

            var shape = textFrame2.Parent as Shape;

            return shape == null ? null : shape.TextFrame.TextRange;
        }
        # endregion

        # region Slide
        public static void ExportSlide(Slide slide, string exportPath)
        {
            slide.Export(exportPath,
                         "PNG",
                         (int) GetDesiredExportWidth(),
                         (int) GetDesiredExportHeight());
        }

        public static void ExportSlide(PowerPointSlide slide, string exportPath)
        {
            ExportSlide(slide.GetNativeSlide(), exportPath);
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
        public static int ConvertColorToRgb(Color argb)
        {
            return (argb.B << 16) | (argb.G << 8) | argb.R;
        }

        public static Color ConvertRgbToColor(int rgb)
        {
            return Color.FromArgb(rgb & 255, (rgb >> 8) & 255, (rgb >> 16) & 255);
        }
        # endregion
        # endregion

        # region Helper Functions
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

        private static void SyncShapeLocation(Shape refShape, Shape candidateShape)
        {
            candidateShape.Left = refShape.Left;
            candidateShape.Top = refShape.Top;
        }

        private static void SyncShapeRotation(Shape refShape, Shape candidateShape)
        {
            candidateShape.Rotation = refShape.Rotation;
        }

        private static void SyncShapeSize(Shape refShape, Shape candidateShape)
        {
            // unlock aspect ratio to enable size tweak
            var candidateLockRatio = candidateShape.LockAspectRatio;

            candidateShape.LockAspectRatio = MsoTriState.msoFalse;

            candidateShape.Width = refShape.Width;
            candidateShape.Height = refShape.Height;

            candidateShape.LockAspectRatio = candidateLockRatio;
        }
        # endregion
    }
}
