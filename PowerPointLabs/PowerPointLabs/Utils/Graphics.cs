using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PPExtraEventHelper;
using PowerPointLabs.Models;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;
using TextFrame2 = Microsoft.Office.Interop.PowerPoint.TextFrame2;
using Drawing = System.Drawing;

namespace PowerPointLabs.Utils
{
    public static class Graphics
    {
#pragma warning disable 0618
        #region Const
        public const float PictureExportingRatio = 96.0f / 72.0f;
        private const float TargetDpi = 96.0f;
        private static float dpiScale = 1.0f;

        // Static initializer to retrieve dpi scale once
        static Graphics()
        {
            using (System.Drawing.Graphics g = System.Drawing.Graphics.FromHwnd(IntPtr.Zero))
            {
                dpiScale = g.DpiX / TargetDpi;
            }
        }
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

        public static bool IsStraightLine(Shape shape)
        {
            return shape.Type == MsoShapeType.msoLine ||
                    (shape.Type == MsoShapeType.msoAutoShape &&
                     shape.AutoShapeType == MsoAutoShapeType.msoShapeMixed &&
                     shape.ConnectorFormat.Type == MsoConnectorType.msoConnectorStraight);
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

        /// <summary>
        /// A better version of SyncShape, but cannot do a partial sync like SyncShape can.
        /// SyncShape cannot operate on Groups and Charts. If those are detected, SyncWholeShape resorts to deleting
        /// candidateShape and replacing it with a copy of refShape instead of syncing.
        /// </summary>
        public static void SyncWholeShape(Shape refShape, ref Shape candidateShape, PowerPointSlide candidateSlide)
        {
            bool succeeded = true;
            try
            {
                SyncShape(refShape, candidateShape);
            }
            catch (UnauthorizedAccessException)
            {
                succeeded = false;
            }
            catch (ArgumentException)
            {
                succeeded = false;
            }
            catch (COMException)
            {
                succeeded = false;
            }
            if (succeeded) return;

            candidateShape.Delete();
            refShape.Copy();
            candidateShape = candidateSlide.Shapes.Paste()[1];
            candidateShape.Name = refShape.Name;
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


            if (pickupShapeFormat)
            {
                refShape.PickUp();
                candidateShape.Apply();
            }

            if ((pickupTextContent || pickupTextFormat) &&
                refShape.HasTextFrame == MsoTriState.msoTrue &&
                candidateShape.HasTextFrame == MsoTriState.msoTrue)
            {
                var refTextRange = refShape.TextFrame2.TextRange;
                var candidateTextRange = candidateShape.TextFrame2.TextRange;

                if (pickupTextContent)
                {
                    candidateTextRange.Text = refTextRange.Text;
                }

                var refParagraphCount = refShape.TextFrame2.TextRange.Paragraphs.Count;
                var candidateParagraphCount = candidateShape.TextFrame2.TextRange.Paragraphs.Count;

                if (refParagraphCount > 0)
                {
                    string originalText = candidateTextRange.Text;
                    SyncTextRange(refTextRange.Paragraphs[refParagraphCount], candidateTextRange);
                    candidateTextRange.Text = originalText;
                }

                for (var i = 1; i <= candidateParagraphCount; i++)
                {
                    var refParagraph = refTextRange.Paragraphs[i <= refParagraphCount ? i : refParagraphCount];
                    var candidateParagraph = candidateTextRange.Paragraphs[i];

                    SyncTextRange(refParagraph, candidateParagraph, pickupTextContent, pickupTextFormat);
                }
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
            bool originallyHadNewLine = candidateTextRange.Text.EndsWith("\r");
            bool lostTheNewLine = false;

            var candidateText = candidateTextRange.Text.TrimEnd('\r');

            if (pickupTextFormat)
            {
                // pick up format using copy-paste, since we could not deep copy the format
                refTextRange.Copy();
                candidateTextRange.PasteSpecial(MsoClipboardFormat.msoClipboardFormatNative);
                lostTheNewLine = !candidateTextRange.Text.EndsWith("\r");
            }

            if (!pickupTextContent)
            {
                candidateTextRange.Text = candidateText;

                // Handling an uncommon edge case. If we are not copying paragraph content, only format,
                // Sometimes (when the reference paragraph doesn't end with a newline), the newline will be lost after copy.
                if (originallyHadNewLine && lostTheNewLine)
                {
                    candidateTextRange.Text = candidateTextRange.Text + "\r";
                }
            }
        }

        /// <summary>
        /// Sort by increasing Z-Order.
        /// (From front to back).
        /// </summary>
        public static void SortByZOrder(List<Shape> shapes)
        {
            shapes.Sort((sh1, sh2) => sh2.ZOrderPosition - sh1.ZOrderPosition);
        }

        /// <summary>
        /// Moves shiftShape forward until it is in front of destinationShape.
        /// (does nothing if already in front)
        /// </summary>
        public static void MoveZUntilInFront(Shape shiftShape, Shape destinationShape)
        {
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

        /// <summary>
        /// Moves shiftShape backward until it is behind destinationShape.
        /// (does nothing if already behind)
        /// </summary>
        public static void MoveZUntilBehind(Shape shiftShape, Shape destinationShape)
        {
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
        /// Moves shiftShape forward/backward until it is just behind destinationShape
        /// </summary>
        public static void MoveZToJustBehind(Shape shiftShape, Shape destinationShape)
        {
            // Step 1: Shift forward until it overshoots destination.
            MoveZUntilInFront(shiftShape, destinationShape);

            // Step 2: Shift backward until it overshoots destination.
            MoveZUntilBehind(shiftShape, destinationShape);
        }

        /// <summary>
        /// Moves shiftShape forward/backward until it is just in front of destinationShape
        /// </summary>
        public static void MoveZToJustInFront(Shape shiftShape, Shape destinationShape)
        {
            // Step 1: Shift backward until it overshoots destination.
            MoveZUntilBehind(shiftShape, destinationShape);

            // Step 2: Shift forward until it overshoots destination.
            MoveZUntilInFront(shiftShape, destinationShape);
        }

        // TODO: This could be an extension method of shape.
        public static bool HasDefaultName(Shape shape)
        {
            var copy = shape.Duplicate()[1];
            bool hasDefaultName = copy.Name != shape.Name;
            copy.Delete();
            return hasDefaultName;
        }

        // TODO: This could be an extension method of shape.
        public static float GetMidpointX(Shape shape)
        {
            return shape.Left + shape.Width / 2;
        }

        // TODO: This could be an extension method of shape.
        public static void SetMidpointX(Shape shape, float value)
        {
            shape.Left = value - shape.Width / 2;
        }

        // TODO: This could be an extension method of shape.
        public static float GetMidpointY(Shape shape)
        {
            return shape.Top + shape.Height / 2;
        }

        // TODO: This could be an extension method of shape.
        public static void SetMidpointY(Shape shape, float value)
        {
            shape.Top = value - shape.Height / 2;
        }

        // TODO: This could be an extension method of shape.
        public static float GetRight(Shape shape)
        {
            return shape.Left + shape.Width;
        }

        // TODO: This could be an extension method of shape.
        public static void SetRight(Shape shape, float value)
        {
            shape.Left = value - shape.Width;
        }

        // TODO: This could be an extension method of shape.
        public static float GetBottom(Shape shape)
        {
            return shape.Top + shape.Height;
        }

        // TODO: This could be an extension method of shape.
        public static void SetBottom(Shape shape, float value)
        {
            shape.Top = value - shape.Height;
        }

        public static float GetScaleWidth(Shape shape)
        {
            float oldWidth = shape.Width;
            shape.ScaleWidth(1, MsoTriState.msoCTrue);
            float scaleFactorWidth = oldWidth / shape.Width;

            shape.ScaleWidth(scaleFactorWidth, MsoTriState.msoCTrue);

            return scaleFactorWidth;
        }

        public static float GetScaleHeight(Shape shape)
        {
            float oldHeight = shape.Height;
            shape.ScaleHeight(1, MsoTriState.msoCTrue);
            float scaleFactorHeight = oldHeight / shape.Height;

            shape.ScaleHeight(scaleFactorHeight, MsoTriState.msoCTrue);

            return scaleFactorHeight;
        }

        // TODO: This could be an extension method of shape.
        /// <summary>
        /// anchorFraction = 0 means left side, anchorFraction = 1 means right side.
        /// </summary>
        public static void SetShapeX(Shape shape, float value, float anchorFraction)
        {
            shape.Left = value - shape.Width * anchorFraction;
        }

        /// <summary>
        /// anchorFraction = 0 means top side, anchorFraction = 1 means bottom side.
        /// </summary>
        public static void SetShapeY(Shape shape, float value, float anchorFraction)
        {
            shape.Top = value - shape.Height * anchorFraction;
        }

        /// <summary>
        /// anchorX and anchorY are between 0 and 1. They represent the pivot to rotate the shape about.
        /// The shape rotates by angle difference angle from its current angle. angle is in degrees.
        /// </summary>
        public static void RotateShapeAboutPivot(Shape shape, float angle, float anchorX, float anchorY)
        {
            double pivotX = shape.Left + anchorX*shape.Width;
            double pivotY = shape.Top + anchorY*shape.Height;
            double midpointX = GetMidpointX(shape);
            double midpointY = GetMidpointY(shape);

            double dx = midpointX - pivotX;
            double dy = midpointY - pivotY;

            double radAngle = angle * Math.PI / 180;
            double newdx = Math.Cos(radAngle) * dx - Math.Sin(radAngle) * dy;
            double newdy = Math.Sin(radAngle) * dx + Math.Cos(radAngle) * dy;

            SetMidpointX(shape, (float)(pivotX + newdx));
            SetMidpointY(shape, (float)(pivotY + newdy));
            shape.Rotation += angle;
        }

        // TODO: This could be an extension method of shape.
        public static string GetText(Shape shape)
        {
            return shape.TextFrame2.TextRange.Text;
        }

        // TODO: This could be an extension method of shape.
        public static void SetText(Shape shape, params string[] lines)
        {
            shape.TextFrame2.TextRange.Text = string.Join("\r", lines);
        }

        // TODO: This could be an extension method of shape.
        public static void SetText(Shape shape, IEnumerable<string> lines)
        {
            shape.TextFrame2.TextRange.Text = string.Join("\r", lines);
        }

        // TODO: This could be an extension method of shape.
        /// <summary>
        /// Get the paragraphs of the shape as a list.
        /// The paragraphs formats can be modified to change the format of the paragraphs in shape.
        /// This list is 0-indexed.
        /// </summary>
        public static List<TextRange2> GetParagraphs(Shape shape)
        {
            return shape.TextFrame2.TextRange.Paragraphs.Cast<TextRange2>().ToList();
        }

        // TODO: This could be an extension method of shape.
        public static bool IsHidden(Shape shape)
        {
            return shape.Visible == MsoTriState.msoFalse;
        }

        public static bool IsAGroup(Shape shape)
        {
            try
            {
                var groupItems = shape.GroupItems;
            }
            catch (UnauthorizedAccessException)
            {
                return false;
            }
            return true;
        }

        public static bool CanAddArrows(Shape shape)
        {
            try
            {
                if (shape.Line.Visible != MsoTriState.msoTrue) return false;
                shape.Line.BeginArrowheadStyle = shape.Line.BeginArrowheadStyle;
                return true;
            }
            catch (ArgumentException)
            {
                return false;
            }
        }

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

        public static PointF GetCenterPoint(Shape s)
        {
            return new PointF(s.Left + s.Width / 2, s.Top + s.Height / 2);
        }

        internal static List<PPShape> SortShapesByLeft(List<PPShape> selectedShapes)
        {
            List<PPShape> shapesToBeSorted = new List<PPShape>();

            for (int i = 0; i < selectedShapes.Count; i++)
            {
                shapesToBeSorted.Add(selectedShapes[i]);
            }

            shapesToBeSorted.Sort(LeftComparator);

            return shapesToBeSorted;
        }

        internal static List<PPShape> SortShapesByTop(List<PPShape> selectedShapes)
        {
            List<PPShape> shapesToBeSorted = new List<PPShape>();

            for (int i = 0; i < selectedShapes.Count; i++)
            {
                shapesToBeSorted.Add(selectedShapes[i]);
            }

            shapesToBeSorted.Sort(TopComparator);

            return shapesToBeSorted;
        }

        private static int LeftComparator(PPShape s1, PPShape s2)
        {
            return s1.VisualLeft.CompareTo(s2.VisualLeft);
        }

        private static int TopComparator(PPShape s1, PPShape s2)
        {
            return s1.VisualTop.CompareTo(s2.VisualTop);
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
            if (inSlide.HasExitAnimation(shape)) return;

            var effectFade = inSlide.GetNativeSlide().TimeLine.MainSequence.AddEffect(shape, MsoAnimEffect.msoAnimEffectAppear,
                MsoAnimateByLevel.msoAnimateLevelNone, MsoAnimTriggerType.msoAnimTriggerWithPrevious, effectStartIndex);
            effectFade.Exit = MsoTriState.msoTrue;
        }

        private static void AddAppearAnimation(Shape shape, PowerPointSlide inSlide, int effectStartIndex)
        {
            if (inSlide.HasEntryAnimation(shape)) return;

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
