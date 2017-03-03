using System;

using PowerPointLabs.CustomControls;

namespace PowerPointLabs.CropLab
{
    internal class CropLabErrorHandler
    {
        private IMessageService View { get; set; }
        private static CropLabErrorHandler _errorHandler;
        
        public const string SelectionTypeShape = "shape";
        public const string SelectionTypePicture = "picture";

        public const int ErrorCodeSelectionIsInvalid = 0;
        public const int ErrorCodeSelectionMustBeShape = 1;
        public const int ErrorCodeSelectionMustBePicture = 2;
        public const int ErrorCodeAspectRatioIsInvalid = 3;
        public const int ErrorCodeSelectionCountZero = 4;
        public const int ErrorCodeUndefined = 5;

        private const string ErrorMessageSelectionIsInvalid = TextCollection.CropLabText.ErrorSelectionIsInvalid;
        private const string ErrorMessageSelectionMustBeShape = TextCollection.CropLabText.ErrorSelectionMustBeShape;
        private const string ErrorMessageSelectionMustBePicture = TextCollection.CropLabText.ErrorSelectionMustBePicture;
        private const string ErrorMessageAspectRatioIsInvalid = TextCollection.CropLabText.ErrorAspectRatioIsInvalid;
        private const string ErrorMessageUndefined = TextCollection.CropLabText.ErrorUndefined;
        private const string ErrorMessageForSelectionCountZero = TextCollection.CropToSlideText.ErrorMessageForSelectionCountZero;

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
                case ErrorCodeSelectionMustBePicture:
                case ErrorCodeSelectionMustBeShape:
                    ShowErrorMessage(errorCode, featureName);
                    break;
                case ErrorCodeSelectionCountZero:
                case ErrorCodeAspectRatioIsInvalid:
                    ShowErrorMessage(errorCode);
                    break;
                default:
                    ShowErrorMessage(errorCode);
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
                    return ErrorMessageSelectionIsInvalid;
                case ErrorCodeSelectionMustBeShape:
                    return ErrorMessageSelectionMustBeShape;
                case ErrorCodeSelectionMustBePicture:
                    return ErrorMessageSelectionMustBePicture;
                case ErrorCodeAspectRatioIsInvalid:
                    return ErrorMessageAspectRatioIsInvalid;
                case ErrorCodeSelectionCountZero:
                    return ErrorMessageForSelectionCountZero;
                default:
                    return ErrorMessageUndefined;
            }
        }
    }
}
