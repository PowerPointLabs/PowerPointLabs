using System;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs
{
    public class CropToAspectRatio
    {
#pragma warning disable 0618
        private const int ErrorCodeForSelectionCountZero = 0;
        private const int ErrorCodeForSelectionNonPicture = 1;
        private const int ErrorCodeForAspectRatioInvalid = 2;

        private const string ErrorMessageForSelectionCountZero = TextCollection.CropToAspectRatioText.ErrorMessageForSelectionCountZero;
        private const string ErrorMessageForSelectionNonPicture = TextCollection.CropToAspectRatioText.ErrorMessageForSelectionNonPicture;
        private const string ErrorMessageForUndefined = TextCollection.CropToAspectRatioText.ErrorMessageForUndefined;

        private const string MessageBoxTitle = "Unable to crop";

        public static PowerPoint.ShapeRange Crop(PowerPoint.Selection selection, float aspectRatioWidth, float aspectRatioHeight, 
                                                 bool handleError = true)
        {
            try
            {
                VerifyIsAspectRatioValid(aspectRatioWidth, aspectRatioHeight);
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

            float aspectRatio = aspectRatioWidth / aspectRatioHeight;
            var croppedShape = Crop(selection.ShapeRange, aspectRatio, handleError: handleError);
            if (croppedShape != null)
            {
                croppedShape.Select();
            }
            return croppedShape;
        }

        public static PowerPoint.ShapeRange Crop(PowerPoint.ShapeRange shapeRange, float aspectRatio, bool handleError = true)
        {
            try
            {
                if (!VerifyIsShapeRangeValid(shapeRange, handleError))
                {
                    return null;
                }
                
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    PowerPoint.ShapeRange origShape = shapeRange[i].Duplicate();
                    origShape.ScaleWidth(1, Microsoft.Office.Core.MsoTriState.msoTrue);
                    origShape.ScaleHeight(1, Microsoft.Office.Core.MsoTriState.msoTrue);
                    float origWidth = origShape.Width;
                    float origHeight = origShape.Height;
                    origShape.Delete();

                    float currentWidth = shapeRange[i].Width - (shapeRange[i].PictureFormat.CropLeft + shapeRange[i].PictureFormat.CropRight) / origWidth;
                    float currentHeight = shapeRange[i].Height - (shapeRange[i].PictureFormat.CropTop + shapeRange[i].PictureFormat.CropBottom) / origHeight;
                    float currentProportions = currentWidth / currentHeight;

                    if (currentProportions > aspectRatio)
                    {
                        // Crop the width
                        float desiredWidth = currentHeight * aspectRatio;
                        float widthToCropEachSide = (currentWidth - desiredWidth) / 2.0f;
                        float widthToCropEachSideRatio = widthToCropEachSide / currentWidth;
                        shapeRange[i].PictureFormat.CropLeft += origWidth * widthToCropEachSideRatio;
                        shapeRange[i].PictureFormat.CropRight += origWidth * widthToCropEachSideRatio;
                    }
                    else if (currentProportions < aspectRatio)
                    {
                        // Crop the height
                        float desiredHeight = currentWidth / aspectRatio;
                        float heightToCropEachSide = (currentHeight - desiredHeight) / 2.0f;
                        float heightToCropEachSideRatio = heightToCropEachSide / currentHeight;
                        shapeRange[i].PictureFormat.CropTop += origHeight * heightToCropEachSideRatio;
                        shapeRange[i].PictureFormat.CropBottom += origHeight * heightToCropEachSideRatio;
                    }
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

        private static void VerifyIsAspectRatioValid(float aspectRatioWidth, float aspectRatioHeight)
        {
            if (aspectRatioWidth <= 0 || aspectRatioHeight <= 0)
            {
                ThrowErrorCode(ErrorCodeForAspectRatioInvalid);
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
