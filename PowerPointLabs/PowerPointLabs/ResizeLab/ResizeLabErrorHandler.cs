using System;

namespace PowerPointLabs.ResizeLab
{
    internal class ResizeLabErrorHandler
    {
        private IResizeLabPane View { get; set; }
        private static ResizeLabErrorHandler _errorHandler;

        public const int ErrorCodeInvalidSelection = 0;

        private const string ErrorMessageInvalidSelection = TextCollection.ResizeLabText.ErrorInvalidSelection;
        private const string ErrorMessageUndefined = TextCollection.ResizeLabText.ErrorUndefined;

        private ResizeLabErrorHandler(IResizeLabPane view = null)
        {
            View = view;
        }

        public static ResizeLabErrorHandler InitializErrorHandler(IResizeLabPane view = null)
        {
            if (_errorHandler == null)
            {
                _errorHandler = new ResizeLabErrorHandler(view);
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
                case ErrorCodeInvalidSelection:
                    return ErrorMessageInvalidSelection;
                default:
                    return ErrorMessageUndefined;
            }
        }
    }
}
