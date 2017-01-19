using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs
{
    public class CropOutPadding
    {
#pragma warning disable 0618
        private const int ErrorCodeForSelectionCountZero = 0;
        private const int ErrorCodeForSelectionNonPicture = 1;

        private const string ErrorMessageForSelectionCountZero = TextCollection.CropOutPaddingText.ErrorMessageForSelectionCountZero;
        private const string ErrorMessageForSelectionNonPicture = TextCollection.CropOutPaddingText.ErrorMessageForSelectionNonPicture;
        private const string ErrorMessageForUndefined = TextCollection.CropOutPaddingText.ErrorMessageForUndefined;

        private const string MessageBoxTitle = "Unable to crop";

        private static readonly string ShapePicture = Path.GetTempPath() + @"\cropoutpaddingtemp.png";

        public static PowerPoint.ShapeRange Crop(PowerPoint.Selection selection, double magnifyRatio = 1.0, bool isInPlace = false,
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

            var croppedShape = Crop(selection.ShapeRange, isInPlace: isInPlace, handleError: handleError);
            if (croppedShape != null)
            {
                croppedShape.Select();
            }
            return croppedShape;
        }

        public static PowerPoint.ShapeRange Crop(PowerPoint.ShapeRange shapeRange, double magnifyRatio = 1.0, bool isInPlace = false,
            bool handleError = true)
        {
            try
            {
                if (!VerifyIsShapeRangeValid(shapeRange, handleError))
                {
                    return null;
                }
                
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    float currentRotation = shapeRange[i].Rotation;
                    float cropLeft = shapeRange[i].PictureFormat.CropLeft;
                    float cropRight = shapeRange[i].PictureFormat.CropRight;
                    float cropTop = shapeRange[i].PictureFormat.CropTop;
                    float cropBottom = shapeRange[i].PictureFormat.CropBottom;
                    
                    shapeRange[i].PictureFormat.CropLeft = 0;
                    shapeRange[i].PictureFormat.CropRight = 0;
                    shapeRange[i].PictureFormat.CropTop = 0;
                    shapeRange[i].PictureFormat.CropBottom = 0;
                    shapeRange[i].Rotation = 0;

                    Utils.Graphics.ExportShape(shapeRange[i], ShapePicture);
                    using (Bitmap shapeImage = new Bitmap(ShapePicture))
                    {
                        PowerPoint.ShapeRange origShape = shapeRange[i].Duplicate();
                        origShape.ScaleWidth(1, Microsoft.Office.Core.MsoTriState.msoTrue);
                        origShape.ScaleHeight(1, Microsoft.Office.Core.MsoTriState.msoTrue);
                        float origWidth = origShape.Width;
                        float origHeight = origShape.Height;
                        origShape.Delete();

                        Rectangle cropRect = GetImageBoundingRect(shapeImage);
                        float cropRatioLeft = cropRect.Left / (float)shapeImage.Width;
                        float cropRatioRight = (shapeImage.Width - cropRect.Width) / (float)shapeImage.Width;
                        float cropRatioTop = cropRect.Top / (float)shapeImage.Height;
                        float cropRatioBottom = (shapeImage.Height - cropRect.Height) / (float)shapeImage.Height;

                        float newCropLeft = origWidth * cropRatioLeft;
                        float newCropRight = origWidth * cropRatioRight;
                        float newCropTop = origHeight * cropRatioTop;
                        float newCropBottom = origHeight * cropRatioBottom;

                        cropLeft = cropLeft < newCropLeft ? newCropLeft : cropLeft;
                        cropRight = cropRight < newCropRight ? newCropRight : cropRight;
                        cropTop = cropTop < newCropTop ? newCropTop : cropTop;
                        cropBottom = cropBottom < newCropBottom ? newCropBottom : cropBottom;
                    }

                    shapeRange[i].Rotation = currentRotation;
                    shapeRange[i].PictureFormat.CropLeft = cropLeft;
                    shapeRange[i].PictureFormat.CropRight = cropRight;
                    shapeRange[i].PictureFormat.CropTop = cropTop;
                    shapeRange[i].PictureFormat.CropBottom = cropBottom;
                }
                return shapeRange;
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

        private static bool IsImageRowTransparent(Bitmap image, int y)
        {
            for (int x = 0; x < image.Width; x++)
            {
                if (image.GetPixel(x, y).A > 0)
                {
                    return false;
                }
            }
            return true;
        }

        private static bool IsImageColumnTransparent(Bitmap image, int x)
        {
            for (int y = 0; y < image.Height; y++)
            {
                if (image.GetPixel(x, y).A > 0)
                {
                    return false;
                }
            }
            return true;
        }

        private static Rectangle GetImageBoundingRect(Bitmap image)
        {
            int left = 0;
            int top = 0;
            int right = 0;
            int bottom = 0;

            // Get left boundary
            for (int x = 0; x < image.Width; x++)
            {
                if (!IsImageColumnTransparent(image, x))
                {
                    left = x;
                    break;
                }
            }

            // Get right boundary
            for (int x = image.Width - 1; x >= 0; x--)
            {
                if (!IsImageColumnTransparent(image, x))
                {
                    right = x;
                    break;
                }
            }

            // Get top boundary
            for (int y = 0; y < image.Height; y++)
            {
                if (!IsImageRowTransparent(image, y))
                {
                    top = y;
                    break;
                }
            }

            // Get bottom boundary
            for (int y = image.Height - 1; y >= 0; y--)
            {
                if (!IsImageRowTransparent(image, y))
                {
                    bottom = y;
                    break;
                }
            }

            Rectangle boundingRect = new Rectangle(left, top, right, bottom);
            return boundingRect;
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
            return shape.Type == Office.MsoShapeType.msoPicture;
        }

        private static void ThrowErrorCode(int typeOfError)
        {
            throw new Exception(typeOfError.ToString(CultureInfo.InvariantCulture));
        }

        private static void IgnoreExceptionThrown(){}

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
