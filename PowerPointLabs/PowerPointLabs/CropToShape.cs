using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs
{
    public class CropToShape
    {
#pragma warning disable 0618
        private const int ErrorCodeForSelectionCountZero = 0;
        private const int ErrorCodeForSelectionNonShape = 1;

        private const string ErrorMessageForSelectionCountZero = TextCollection.CropToShapeText.ErrorMessageForSelectionCountZero;
        private const string ErrorMessageForSelectionNonShape = TextCollection.CropToShapeText.ErrorMessageForSelectionNonShape;
        private const string ErrorMessageForUndefined = TextCollection.CropToShapeText.ErrorMessageForUndefined;

        private const string MessageBoxTitle = "Unable to crop";

        private static float currentMagnifyRatio = 1.0f;
        private const float MinMagnifyRatio = 0.1f;
        private const float MaxMagnifyRatio = 2.0f; // we don't want to export too large resolution and load for too long

        private static readonly string SlidePicture = Path.GetTempPath() + @"\slide.png";
        private static readonly string FillInBackgroundPicture = Path.GetTempPath() + @"\currentFillInBg.png";

        public static PowerPoint.Shape Crop(PowerPoint.Selection selection, float magnifyRatio = 1.0f, bool isInPlace = false,
                                            bool handleError = true)
        {
            try
            {
                VerifyIsSelectionValid(selection);
            }
            catch (Exception e)
            {
                if (handleError)
                {
                    ProcessErrorMessage(e);
                    return null;
                }

                throw;
            }
            
            var croppedShape = Crop(selection.ShapeRange, magnifyRatio: magnifyRatio, isInPlace: isInPlace, handleError: handleError);
            if (croppedShape != null)
            {
                croppedShape.Select();
            }

            return croppedShape;
        }

        public static PowerPoint.Shape Crop(PowerPoint.ShapeRange shapeRange, float magnifyRatio = 1.0f, bool isInPlace = false,
            bool handleError = true)
        {
            try
            {
                if (!VerifyIsShapeRangeValid(shapeRange, handleError)) return null;
                
                var hasManyShapes = shapeRange.Count > 1;
                var shape = hasManyShapes ? shapeRange.Group() : shapeRange[1];
                var left = shape.Left;
                var top = shape.Top;
                shape.Cut();
                shapeRange = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes.Paste();
                shapeRange.Left = left;
                shapeRange.Top = top;
                if (hasManyShapes)
                {
                    shapeRange = shapeRange.Ungroup();
                }

                SetMagnifyRatio(magnifyRatio);
                TakeScreenshotProxy(shapeRange);

                var ungroupedRange = UngroupAllForShapeRange(shapeRange);
                var shapeNames = new string[ungroupedRange.Count];

                for (int i = 1; i <= ungroupedRange.Count; i++)
                {
                    var filledShape = FillInShapeWithImage(SlidePicture, ungroupedRange[i], isInPlace);
                    shapeNames[i - 1] = filledShape.Name;
                }
                
                var croppedRange = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes.Range(shapeNames);
                var croppedShape = (croppedRange.Count == 1) ? croppedRange[1] : croppedRange.Group();
                
                return croppedShape;
            }
            catch (Exception e)
            {
                if (handleError)
                {
                    ProcessErrorMessage(e);
                    return null;
                }

                throw;
            }
        }

        public static PowerPoint.Shape FillInShapeWithImage(string imageFile, PowerPoint.Shape shape, bool isInPlace = false)
        {
            CreateFillInBackgroundForShape(imageFile, shape);
            shape.Fill.UserPicture(FillInBackgroundPicture);

            shape.Line.Visible = Office.MsoTriState.msoFalse;

            if (isInPlace)
            {
                return shape;
            }

            shape.Copy();
            var shapeToReturn = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes.Paste()[1];
            shape.Delete();
            return shapeToReturn;
        }

        public static Bitmap KiCut(Bitmap original, float startX, float startY, float width, float height,
                            float magnifyRatio = 1.0f)
        {
            if (original == null) return null;
            try
            {
                var newX = startX * magnifyRatio;
                var newY = startY * magnifyRatio;
                var newWidth = width * magnifyRatio;
                var newHeight = height * magnifyRatio;

                var outputImage = new Bitmap((int)newWidth, (int)newHeight, PixelFormat.Format32bppArgb);

                var inputGraphics = Graphics.FromImage(outputImage);
                inputGraphics.DrawImage(original,
                    new Rectangle(0, 0, (int)newWidth, (int)newHeight),
                    new Rectangle((int)newX, (int)newY, (int)newWidth, (int)newHeight),
                    GraphicsUnit.Pixel);
                inputGraphics.Dispose();

                return outputImage;
            }
            catch
            {
                return null;
            }
        }

        public static string GetErrorMessageForErrorCode(string errorCode)
        {
            var errorCodeInteger = -1;
            try
            {
                errorCodeInteger = Int32.Parse(errorCode);
            }
            catch
            {
                IgnoreExceptionThrown();
            }
            switch (errorCodeInteger)
            {
                case ErrorCodeForSelectionCountZero:
                    return ErrorMessageForSelectionCountZero;
                case ErrorCodeForSelectionNonShape:
                    return ErrorMessageForSelectionNonShape;
                default:
                    return ErrorMessageForUndefined;
            }
        }

        public static Bitmap GetCutOutShapeMenuImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.CutOutShapeMenu);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetCutOutShapeMenuImage");
                throw;
            }
        }

        private static void CreateFillInBackgroundForShape(string imageFile, PowerPoint.Shape shape)
        {
            using (var slideImage = (Bitmap)Image.FromFile(imageFile))
            {
                if (shape.Rotation == 0)
                {
                    CreateFillInBackground(shape, slideImage);
                }
                else
                {
                    CreateRotatedFillInBackground(shape, slideImage);
                }
            }
        }

        private static void CreateFillInBackground(PowerPoint.Shape shape, Bitmap slideImage)
        {
            var croppedImage = KiCut(slideImage,
                shape.Left * Utils.Graphics.PictureExportingRatio,
                shape.Top * Utils.Graphics.PictureExportingRatio,
                shape.Width * Utils.Graphics.PictureExportingRatio,
                shape.Height * Utils.Graphics.PictureExportingRatio,
                currentMagnifyRatio);
            croppedImage.Save(FillInBackgroundPicture, ImageFormat.Png);
        }

        private static void CreateRotatedFillInBackground(PowerPoint.Shape shape, Bitmap slideImage)
        {
            var rotatedShape = new Utils.PPShape(shape, false);
            var topLeftPoint = new PointF(rotatedShape.ActualTopLeft.X * Utils.Graphics.PictureExportingRatio,
                rotatedShape.ActualTopLeft.Y * Utils.Graphics.PictureExportingRatio);

            Bitmap rotatedImage = new Bitmap(slideImage.Width, slideImage.Height);

            using (Graphics g = Graphics.FromImage(rotatedImage))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;

                using (System.Drawing.Drawing2D.Matrix mat = new System.Drawing.Drawing2D.Matrix())
                {
                    mat.Translate(-topLeftPoint.X, -topLeftPoint.Y);
                    mat.RotateAt(-shape.Rotation, topLeftPoint);

                    g.Transform = mat;
                    g.DrawImage(slideImage, new Rectangle(0, 0, slideImage.Width, slideImage.Height));
                }
            }

            var magnifiedImage = KiCut(rotatedImage, 0, 0, 
                                        shape.Width * Utils.Graphics.PictureExportingRatio,
                                        shape.Height * Utils.Graphics.PictureExportingRatio, 
                                        currentMagnifyRatio);
            magnifiedImage.Save(FillInBackgroundPicture, ImageFormat.Png);
        }

        private static void TakeScreenshotProxy(PowerPoint.ShapeRange shapeRange)
        {
            shapeRange.Visible = Office.MsoTriState.msoFalse;
            Utils.Graphics.ExportSlide(PowerPointCurrentPresentationInfo.CurrentSlide, SlidePicture, currentMagnifyRatio);
            shapeRange.Visible = Office.MsoTriState.msoTrue;
        }

        private static void SetMagnifyRatio(float magnifyRatio)
        {
            if (magnifyRatio > MaxMagnifyRatio)
            {
                currentMagnifyRatio = MaxMagnifyRatio;
            }
            else if (magnifyRatio < MinMagnifyRatio)
            {
                currentMagnifyRatio = MinMagnifyRatio;
            }
            else
            {
                currentMagnifyRatio = magnifyRatio;
            }
        }

        private static PowerPoint.ShapeRange UngroupAllForShapeRange(PowerPoint.ShapeRange range)
        {
            var ungroupedShapeNames = new List<string>();
            var queue = new Queue<PowerPoint.Shape>();

            foreach (var item in range)
            {
                queue.Enqueue(item as PowerPoint.Shape);
            }
            while (queue.Count != 0)
            {
                var shape = queue.Dequeue();
                if (shape.Type == Office.MsoShapeType.msoGroup)
                {
                    var subRange = shape.Ungroup();
                    foreach (var item in subRange)
                    {
                        queue.Enqueue(item as PowerPoint.Shape);
                    }
                }
                else if (!IsShape(shape))
                {
                    ThrowErrorCode(ErrorCodeForSelectionNonShape);
                }
                else
                {
                    ungroupedShapeNames.Add(shape.Name);
                }
            }
            return PowerPointCurrentPresentationInfo.CurrentSlide.Shapes.Range(ungroupedShapeNames.ToArray());
        }

        private static bool IsShapeForSelection(PowerPoint.ShapeRange shapeRange)
        {
            return (from PowerPoint.Shape shape in shapeRange select shape).All(IsShape);
        }

        private static bool IsShape(PowerPoint.Shape shape)
        {
            return shape.Type == Office.MsoShapeType.msoAutoShape
                || shape.Type == Office.MsoShapeType.msoFreeform
                || shape.Type == Office.MsoShapeType.msoGroup;
        }

        private static void ThrowErrorCode(int typeOfError)
        {
            throw new Exception(typeOfError.ToString(CultureInfo.InvariantCulture));
        }

        private static void IgnoreExceptionThrown() { }

        private static void ProcessErrorMessage(Exception e)
        {
            //This method prompts the error message to user. If it has an unrecognised error code,
            //an alternative message window with erro trace stack pops up and prompts the user to
            //send the trace stack to the developer team.
            var errMessage = GetErrorMessageForErrorCode(e.Message);
            if (!string.Equals(errMessage, ErrorMessageForUndefined, StringComparison.Ordinal))
            {
                MessageBox.Show(errMessage, MessageBoxTitle);
            }
            else
            {
                Views.ErrorDialogWrapper.ShowDialog(MessageBoxTitle, e.Message, e);
            }
        }

        private static bool VerifyIsShapeRangeValid(PowerPoint.ShapeRange shapeRange, bool handleError)
        {
            try
            {
                if (shapeRange.Count < 1)
                {
                    ThrowErrorCode(ErrorCodeForSelectionCountZero);
                }

                if (!IsShapeForSelection(shapeRange))
                {
                    ThrowErrorCode(ErrorCodeForSelectionNonShape);
                }

                return true;
            }
            catch (Exception e)
            {
                if (handleError)
                {
                    ProcessErrorMessage(e);
                    return false;
                }

                throw;
            }
        }

        private static void VerifyIsSelectionValid(PowerPoint.Selection selection)
        {
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                ThrowErrorCode(ErrorCodeForSelectionCountZero);
            }
        }
    }
}
