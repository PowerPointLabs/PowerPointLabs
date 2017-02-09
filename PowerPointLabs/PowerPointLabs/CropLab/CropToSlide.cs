using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

using PowerPointLabs.Models;

using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.CropLab
{
    public class CropToSlide
    {
#pragma warning disable 0618
        private const int ErrorCodeForSelectionCountZero = 0;
        private const int ErrorCodeForSelectionNonPicture = 1;

        private const string ErrorMessageForSelectionCountZero = TextCollection.CropToSlideText.ErrorMessageForSelectionCountZero;
        private const string ErrorMessageForSelectionNonPicture = TextCollection.CropToSlideText.ErrorMessageForSelectionNonPicture;
        private const string ErrorMessageForUndefined = TextCollection.CropToSlideText.ErrorMessageForUndefined;

        private const string MessageBoxTitle = "Unable to crop";

        private static readonly string ShapePicture = Path.GetTempPath() + @"\shape.png";

        public static void Crop(PowerPoint.Selection selection, float slideWidth, float slideHeight,
            double magnifyRatio = 1.0, bool isInPlace = false, bool handleError = true)
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
                    return;
                }

                throw;
            }

            Crop(selection.ShapeRange, slideWidth, slideHeight, isInPlace: isInPlace, handleError: handleError);
        }

        public static void Crop(PowerPoint.ShapeRange shapeRange, float slideWidth, float slideHeight,
            double magnifyRatio = 1.0, bool isInPlace = false, bool handleError = true)
        {
            try
            {
                if (!VerifyIsShapeRangeValid(shapeRange, handleError)) return;
                foreach (PowerPoint.Shape shape in shapeRange)
                {
                    PowerPoint.Shape toRotate = shape;
                    if (shape.Rotation != 0)
                    {
                        RectangleF location = GetAbsoluteBounds(shape);
                        Utils.Graphics.ExportShape(shape, ShapePicture);
                        var newShape = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes.AddPicture(ShapePicture,
                            Office.MsoTriState.msoFalse,
                            Office.MsoTriState.msoTrue,
                            location.Left, location.Top, location.Width, location.Height);
                        toRotate = newShape;
                        toRotate.Name = shape.Name;
                        shape.Delete();

                    }
                    RectangleF cropArea = GetCropArea(toRotate, slideWidth, slideHeight);
                    toRotate.PictureFormat.Crop.ShapeHeight = cropArea.Height;
                    toRotate.PictureFormat.Crop.ShapeWidth = cropArea.Width;
                    toRotate.PictureFormat.Crop.ShapeLeft = cropArea.Left;
                    toRotate.PictureFormat.Crop.ShapeTop = cropArea.Top;
                }
            }
            catch (Exception e)
            {
                if (handleError)
                {
                    ProcessErrorMessage(e);
                    return;
                }

                throw;
            }

        }

        private static RectangleF GetAbsoluteBounds(PowerPoint.Shape shape)
        {
            float rotation = (float)Utils.Graphics.DegreeToRadian(shape.Rotation);
            PointF[] corners = new PointF[]
            {
                new PointF(-shape.Width / 2, -shape.Height / 2),
                new PointF(shape.Width / 2, -shape.Height / 2),
                new PointF(-shape.Width / 2, shape.Height / 2),
                new PointF(shape.Width / 2, shape.Height / 2)
            };
            float minX = float.MaxValue;
            float minY = float.MaxValue;
            float maxX = float.MinValue;
            float maxY = float.MinValue;
            for (int i = 0; i < corners.Length; i++)
            {
                PointF rotated = RotatePoint(corners[i], rotation);
                minX = Math.Min(rotated.X, minX);
                minY = Math.Min(rotated.Y, minY);
                maxX = Math.Max(rotated.X, maxX);
                maxY = Math.Max(rotated.Y, maxY);
            }
            return new RectangleF(shape.Left + shape.Width / 2 + minX, shape.Top + shape.Height / 2 + minY,
                                  maxX - minX, maxY - minY);
        }

        private static PointF RotatePoint(PointF point, float theta)
        {
            return new PointF((float)(point.X * Math.Cos(theta) - point.Y * Math.Sin(theta)),
                            (float)(point.X * Math.Sin(theta) + point.Y * Math.Cos(theta)));
        }
        private static RectangleF GetCropArea(PowerPoint.Shape shape, float slideWidth, float slideHeight)
        {
            float cropTop = Math.Max(0, shape.Top);
            float cropLeft = Math.Max(0, shape.Left);
            float cropHeight = shape.Height - Math.Max(0, -shape.Top);
            float cropWidth = shape.Width - Math.Max(0, -shape.Left);

            cropHeight = Math.Min(slideHeight - cropTop, cropHeight);
            cropWidth = Math.Min(slideWidth - cropLeft, cropWidth);

            return new RectangleF(cropLeft, cropTop, cropWidth, cropHeight);
        }

        private static bool VerifyIsShapeRangeValid(PowerPoint.ShapeRange shapeRange, bool handleError)
        {
            try
            {
                if (shapeRange.Count < 1)
                {
                    ThrowErrorCode(ErrorCodeForSelectionCountZero);
                }

                if (!IsPictureForSelection(shapeRange))
                {
                    ThrowErrorCode(ErrorCodeForSelectionNonPicture);
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

        private static bool IsPictureForSelection(PowerPoint.ShapeRange shapeRange)
        {
            return (from PowerPoint.Shape shape in shapeRange select shape).All(IsPicture);
        }

        private static bool IsPicture(PowerPoint.Shape shape)
        {
            return shape.Type == Office.MsoShapeType.msoPicture ||
                   shape.Type == Office.MsoShapeType.msoLinkedPicture;
        }

        private static void ThrowErrorCode(int typeOfError)
        {
            throw new Exception(typeOfError.ToString(CultureInfo.InvariantCulture));
        }

        private static void IgnoreExceptionThrown() { }

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
                case ErrorCodeForSelectionNonPicture:
                    return ErrorMessageForSelectionNonPicture;
                default:
                    return ErrorMessageForUndefined;
            }
        }

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
        
    }
}
