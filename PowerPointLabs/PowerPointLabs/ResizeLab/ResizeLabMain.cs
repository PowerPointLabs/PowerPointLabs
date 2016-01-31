using System;
using System.Globalization;
using System.Linq;
using System.Windows;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    internal static partial class ResizeLabMain
    {
        private const string MessageBoxTitle = "Error during Resizing";

        private const int ErrorCodeNoSelection = 0;
        private const int ErrorCodeNonShapeSelection = 1;

        private const string ErrorMessageNoSelection = TextCollection.ResizeLabText.ErrorNoSelection;
        private const string ErrorMessageNonShapeSelection = TextCollection.ResizeLabText.ErrorNonShapeSelection;
        private const string ErrorMessageUndefined = TextCollection.ResizeLabText.ErrorUndefined;        

        #region Error Message

        private static void ThrowErrorCode(int errorType)
        {
            throw new Exception(errorType.ToString(CultureInfo.InvariantCulture));
        }

        private static void ProcessErrorMessage(Exception e)
        {
            var errorMessage = GetErrorMessage(e.Message);
            if (!string.Equals(errorMessage, ErrorMessageUndefined, StringComparison.Ordinal))
            {
                MessageBox.Show(errorMessage, MessageBoxTitle);
            }
            else
            {
                Views.ErrorDialogWrapper.ShowDialog(MessageBoxTitle, e.Message, e);
            }
        }

        private static string GetErrorMessage(string errorCode)
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
                case ErrorCodeNoSelection:
                    return ErrorMessageNoSelection;
                case ErrorCodeNonShapeSelection:
                    return ErrorMessageNonShapeSelection;
                default:
                    return ErrorMessageUndefined;
            }
        }

        private static void IgnoreExceptionThrown() { }

        #endregion
    }
}
