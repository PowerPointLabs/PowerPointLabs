using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs
{
    public class CropToSame
    {
#pragma warning disable 0618
        private const int ErrorCodeForSelectionCountZero = 0;
        private const int ErrorCodeForSelectionNonPicture = 1;

        private const string ErrorMessageForSelectionCountZero = TextCollection.CropToSameText.ErrorMessageForSelectionCountZero;
        private const string ErrorMessageForSelectionNonPicture = TextCollection.CropToSameText.ErrorMessageForSelectionNonPicture;
        private const string ErrorMessageForUndefined = TextCollection.CropToSameText.ErrorMessageForUndefined;

        private const string MessageBoxTitle = "Unable to crop";

        private static readonly string ShapePicture = Path.GetTempPath() + @"\shape.png";
        

        public static void StartCropToSame(PowerPoint.Selection selection, bool handleError = true)
        {
            try
            {
                VerifyIsSelectionValid(selection);
                if (!VerifyIsShapeRangeValid(selection.ShapeRange, handleError)) return;
                var shapes = selection.ShapeRange;
                var refShape = shapes[1];
                float refScaleWidth = PowerPointLabs.Utils.Graphics.GetScaleWidth(refShape);
                float refScaleHeight = PowerPointLabs.Utils.Graphics.GetScaleHeight(refShape);
                //MessageBox.Show(shapes[1].PictureFormat.CropTop.ToString() + " " + shapes[1].PictureFormat.CropBottom.ToString() + " " + shapes[1].Height.ToString() + " " + );
                //refShape.ScaleHeight(0.5F, Microsoft.Office.Core.MsoTriState.msoFalse);
                float epsilon = 0.001F;
                for (int i = 2; i <= shapes.Count; i++)
                {
                    
                    float scaleWidth = PowerPointLabs.Utils.Graphics.GetScaleWidth(shapes[i]);
                    float scaleHeight = PowerPointLabs.Utils.Graphics.GetScaleHeight(shapes[i]);
                    float heightToCrop = shapes[i].Height - refShape.Height;
                    float widthToCrop = shapes[i].Width - refShape.Width;

                    float cropTop = Math.Max(shapes[1].PictureFormat.CropTop, epsilon);
                    float cropBottom = Math.Max(shapes[1].PictureFormat.CropBottom, epsilon);
                    float cropLeft = Math.Max(shapes[1].PictureFormat.CropLeft, epsilon);
                    float cropRight = Math.Max(shapes[1].PictureFormat.CropRight, epsilon);

                    float refShapeCroppedHeight = cropTop + cropBottom;
                    float refShapeCroppedWidth = cropLeft + cropRight;
                    
                    shapes[i].PictureFormat.CropTop = Math.Max(0, heightToCrop * cropTop / refShapeCroppedHeight / scaleHeight);
                    shapes[i].PictureFormat.CropLeft = Math.Max(0, widthToCrop * cropLeft / refShapeCroppedWidth / scaleWidth);
                    shapes[i].PictureFormat.CropRight = Math.Max(0, widthToCrop * cropRight / refShapeCroppedWidth / scaleWidth);
                    shapes[i].PictureFormat.CropBottom = Math.Max(0, heightToCrop * cropBottom / refShapeCroppedHeight / scaleHeight);
                    
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
