using System.Linq;
using System.Text.RegularExpressions;

using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.CropLab
{
    internal class CropToAspectRatio
    {
        private static float aspectRatioWidth = 0.0f;
        private static float aspectRatioHeight = 0.0f;

        public static PowerPoint.ShapeRange Crop(PowerPoint.Selection selection, string aspectRatioRawString, 
                                                          CropLabErrorHandler errorHandler = null)
        {
            if (!VerifyIsAspectRatioValid(aspectRatioRawString, errorHandler) ||
                !VerifyIsSelectionValid(selection, errorHandler))
            {
                return null;
            }

            float aspectRatio = aspectRatioWidth / aspectRatioHeight;
            var croppedShape = Crop(selection.ShapeRange, aspectRatio, errorHandler);
            if (croppedShape != null)
            {
                croppedShape.Select();
            }
            return croppedShape;
        }

        public static PowerPoint.ShapeRange Crop(PowerPoint.ShapeRange shapeRange, float aspectRatio, CropLabErrorHandler errorHandler = null)
        {
            if (!VerifyIsShapeRangeValid(shapeRange, errorHandler))
            {
                return null;
            }

            for (int i = 1; i <= shapeRange.Count; i++)
            {
                PowerPoint.ShapeRange origShape = shapeRange[i].Duplicate();
                origShape.ScaleWidth(1, Office.MsoTriState.msoTrue);
                origShape.ScaleHeight(1, Office.MsoTriState.msoTrue);
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

        private static bool VerifyIsSelectionValid(PowerPoint.Selection selection, CropLabErrorHandler errorHandler)
        {
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                HandleErrorCode(CropLabErrorHandler.ErrorCodeSelectionIsInvalid, errorHandler);
                return false;
            }

            return true;
        }

        private static bool VerifyIsShapeRangeValid(PowerPoint.ShapeRange shapeRange, CropLabErrorHandler errorHandler)
        {
            if (shapeRange.Count < 1)
            {
                HandleErrorCode(CropLabErrorHandler.ErrorCodeSelectionIsInvalid, errorHandler);
                return false;
            }

            if (!IsPictureForSelection(shapeRange))
            {
                HandleErrorCode(CropLabErrorHandler.ErrorCodeSelectionMustBePicture, errorHandler);
                return false;
            }

            return true;
        }

        private static bool VerifyIsAspectRatioValid(string aspectRatioString, CropLabErrorHandler errorHandler)
        {
            string pattern = @"(\d+):(\d+)";
            Match matches = Regex.Match(aspectRatioString, pattern);
            if (!matches.Success)
            {
                HandleErrorCode(CropLabErrorHandler.ErrorCodeAspectRatioIsInvalid, errorHandler);
                return false;
            }

            if (!float.TryParse(matches.Groups[1].Value, out aspectRatioWidth) ||
                !float.TryParse(matches.Groups[2].Value, out aspectRatioHeight))
            {
                HandleErrorCode(CropLabErrorHandler.ErrorCodeAspectRatioIsInvalid, errorHandler);
                return false;
            }
            
            if (aspectRatioWidth <= 0.0f || aspectRatioHeight <= 0.0f)
            {
                HandleErrorCode(CropLabErrorHandler.ErrorCodeAspectRatioIsInvalid, errorHandler);
                return false;
            }

            return true;
        }

        private static bool IsPictureForSelection(PowerPoint.ShapeRange shapeRange)
        {
            return (from PowerPoint.Shape shape in shapeRange select shape).All(IsPicture);
        }

        private static bool IsPicture(PowerPoint.Shape shape)
        {
            return shape.Type == Office.MsoShapeType.msoPicture;
        }

        private static void HandleErrorCode(int errorCode, CropLabErrorHandler errorHandler)
        {
            if (errorHandler == null)
            {
                return;
            }

            switch (errorCode)
            {
                case CropLabErrorHandler.ErrorCodeSelectionIsInvalid:
                    errorHandler.ProcessErrorCode(errorCode, "Crop To Aspect Ratio", "1", "picture");
                    break;
                case CropLabErrorHandler.ErrorCodeSelectionMustBePicture:
                    errorHandler.ProcessErrorCode(errorCode, "Crop To Aspect Ratio");
                    break;
                case CropLabErrorHandler.ErrorCodeAspectRatioIsInvalid:
                    errorHandler.ProcessErrorCode(errorCode);
                    break;
                default:
                    errorHandler.ProcessErrorCode(errorCode);
                    break;
            }
        }
    }
}
