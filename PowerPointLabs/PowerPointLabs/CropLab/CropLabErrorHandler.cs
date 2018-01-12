using System;

using PowerPointLabs.CustomControls;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.CropLab
{
    internal class CropLabErrorHandler
    {
        public const string SelectionTypeShape = "shape";
        public const string SelectionTypePicture = "picture";
        public const string SelectionTypeShapeOrPicture = "shape or picture";

        public const int ErrorCodeSelectionIsInvalid = 0;
        public const int ErrorCodeSelectionMustBeShape = 1;
        public const int ErrorCodeSelectionMustBePicture = 2;
        public const int ErrorCodeAspectRatioIsInvalid = 3;
        public const int ErrorCodeSelectionCountZero = 4;
        public const int ErrorCodeUndefined = 5;
        public const int ErrorCodeSelectionMustBeShapeOrPicture = 6;
        public const int ErrorCodeNoShapeOverBoundary = 7;
        public const int ErrorCodeNoDimensionCropped = 8;
        public const int ErrorCodeNoPaddingCropped = 9;
        public const int ErrorCodeNoAspectRatioCropped = 10;

        private IMessageService View { get; set; }
        private static CropLabErrorHandler _errorHandler;

        private CropLabErrorHandler(IMessageService view = null)
        {
            View = view;
        }

        public static CropLabErrorHandler InitializeErrorHandler(IMessageService view = null)
        {
            if (_errorHandler == null)
            {
                _errorHandler = new CropLabErrorHandler(view);
            }
            else if (view != null) // Allow the view to change
            {
                _errorHandler.View = view;
            }
            return _errorHandler;
        }

        public void ProcessErrorCode(int errorCode, string featureName, string validSelectionType = "", int validSelectionMinCount = -1)
        {
            switch (errorCode)
            {
                case ErrorCodeSelectionIsInvalid:
                    if (validSelectionMinCount != 1)
                    {
                        validSelectionType += "s";
                    }
                    ShowErrorMessage(errorCode, featureName, validSelectionMinCount.ToString(), validSelectionType);
                    break;
                case ErrorCodeSelectionMustBeShapeOrPicture:
                case ErrorCodeSelectionMustBePicture:
                case ErrorCodeSelectionMustBeShape:
                case ErrorCodeSelectionCountZero:
                case ErrorCodeAspectRatioIsInvalid:
                case ErrorCodeNoShapeOverBoundary:
                case ErrorCodeNoDimensionCropped:
                case ErrorCodeNoPaddingCropped:
                case ErrorCodeNoAspectRatioCropped:
                default:
                    ShowErrorMessage(errorCode, featureName);
                    break;
            }
        }

        public void ProcessException(Exception e, string message)
        {
            if (View == null) // Nothing to display on
            {
                return;
            }
            View.ShowErrorMessageBox(message, e);
        }

        /// <summary>
        /// Store error code in the culture info.
        /// </summary>
        /// <param name="errorType"></param>
        /// <param name="optionalParameters"></param>
        private void ShowErrorMessage(int errorType, params string[] optionalParameters)
        {
            if (View == null) // Nothing to display on
            {
                return;
            }
            var errorMsg = string.Format(GetErrorMessage(errorType), optionalParameters);
            View.ShowErrorMessageBox(errorMsg);
        }

        /// <summary>
        /// Get error message corresponds to the error code.
        /// </summary>
        /// <param name="errorCode"></param>
        /// <returns></returns>
        private string GetErrorMessage(int errorCode)
        {   
            switch (errorCode)
            {
                case ErrorCodeSelectionIsInvalid:
                    return CropLabText.ErrorSelectionIsInvalid;
                case ErrorCodeSelectionMustBeShape:
                    return CropLabText.ErrorSelectionMustBeShape;
                case ErrorCodeSelectionMustBeShapeOrPicture:
                    return CropLabText.ErrorSelectionMustBeShapeOrPicture;
                case ErrorCodeSelectionMustBePicture:
                    return CropLabText.ErrorSelectionMustBePicture;
                case ErrorCodeAspectRatioIsInvalid:
                    return CropLabText.ErrorAspectRatioIsInvalid;
                case ErrorCodeSelectionCountZero:
                    return CropLabText.ErrorSelectionCountZero;
                case ErrorCodeNoShapeOverBoundary:
                    return CropLabText.ErrorNoShapeOverBoundary;
                case ErrorCodeNoDimensionCropped:
                    return CropLabText.ErrorNoDimensionCropped;
                case ErrorCodeNoPaddingCropped:
                    return CropLabText.ErrorNoPaddingCropped;
                case ErrorCodeNoAspectRatioCropped:
                    return CropLabText.ErrorNoAspectRatioCropped;
                default:
                    return CropLabText.ErrorUndefined;
            }
        }
    }
}
