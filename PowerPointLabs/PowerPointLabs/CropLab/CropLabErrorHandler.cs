namespace PowerPointLabs.CropLab
{
    internal class CropLabErrorHandler
    {
        private CropLabUIControl View { get; set; }
        private static CropLabErrorHandler _errorHandler;
        
        public const int ErrorCodeSelectionIsInvalid = 0;
        public const int ErrorCodeSelectionMustBeShape = 1;
        public const int ErrorCodeSelectionMustBePicture = 2;
        public const int ErrorCodeAspectRatioIsInvalid = 3;

        private const string ErrorMessageSelectionIsInvalid = TextCollection.CropLabText.ErrorSelectionIsInvalid;
        private const string ErrorMessageSelectionMustBeShape = TextCollection.CropLabText.ErrorSelectionMustBeShape;
        private const string ErrorMessageSelectionMustBePicture = TextCollection.CropLabText.ErrorSelectionMustBePicture;
        private const string ErrorMessageAspectRatioIsInvalid = TextCollection.CropLabText.ErrorAspectRatioIsInvalid;
        private const string ErrorMessageUndefined = TextCollection.CropLabText.ErrorUndefined;

        private CropLabErrorHandler(CropLabUIControl view = null)
        {
            View = view;
        }

        public static CropLabErrorHandler InitializeErrorHandler(CropLabUIControl view = null)
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

        /// <summary>
        /// Store error code in the culture info.
        /// </summary>
        /// <param name="errorType"></param>
        /// <param name="optionalParameters"></param>
        public void ProcessErrorCode(int errorType, params string[] optionalParameters)
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
                default:
                    return ErrorMessageUndefined;
            }
        }
    }
}
