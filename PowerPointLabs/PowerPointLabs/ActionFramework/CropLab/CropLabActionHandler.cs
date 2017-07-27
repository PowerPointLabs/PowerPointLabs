using System.Text.RegularExpressions;

using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.CropLab;

namespace PowerPointLabs.ActionFramework.CropLab
{
    abstract class CropLabActionHandler : ActionHandler
    {
        protected static void HandleErrorCode(int errorCode, string featureName, CropLabErrorHandler errorHandler)
        {
            if (errorHandler == null)
            {
                return;
            }
            errorHandler.ProcessErrorCode(errorCode, featureName);
        }

        protected static void HandleInvalidSelectionError(int errorCode, string featureName, string validSelectionType, int validSelectionMinCount, CropLabErrorHandler errorHandler)
        {
            if (errorHandler == null)
            {
                return;
            }
            errorHandler.ProcessErrorCode(errorCode, featureName, validSelectionType, validSelectionMinCount);
        }

        protected static void HandleCropLabException(CropLabException e, string featureName, CropLabErrorHandler errorHandler)
        {
            if (errorHandler == null)
            {
                return;
            }

            if (e.Message.Equals(CropLabErrorHandler.ErrorCodeNoAspectRatioCropped.ToString()))
            {
                HandleErrorCode(CropLabErrorHandler.ErrorCodeNoAspectRatioCropped, featureName, errorHandler);
            }
            else if (e.Message.Equals(CropLabErrorHandler.ErrorCodeNoPaddingCropped.ToString()))
            {
                HandleErrorCode(CropLabErrorHandler.ErrorCodeNoPaddingCropped, featureName, errorHandler);
            }
        }

        protected static bool TryParseAspectRatio(string aspectRatioString, out float aspectRatioWidth, out float aspectRatioHeight)
        {
            aspectRatioWidth = 0.0f;
            aspectRatioHeight = 0.0f;

            string pattern = @"(\d*\.?\d*):(\d*\.?\d*)";
            Match matches = Regex.Match(aspectRatioString, pattern);
            if (!matches.Success)
            {
                return false;
            }

            if (!float.TryParse(matches.Groups[1].Value, out aspectRatioWidth) ||
                !float.TryParse(matches.Groups[2].Value, out aspectRatioHeight))
            {
                return false;
            }

            if (aspectRatioWidth <= 0.0f || aspectRatioHeight <= 0.0f)
            {
                return false;
            }

            return true;
        }
    }
}
