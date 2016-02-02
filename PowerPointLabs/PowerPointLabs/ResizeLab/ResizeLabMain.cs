using System;
using System.Globalization;
using System.Linq;
using System.Windows;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    internal static partial class ResizeLabMain
    {
        private const string MessageBoxTitle = "Error during Resizing";

        private const int ErrorCodeNoSelection = 0;
        private const int ErrorCodeFewerThanTwoSelection = 1;

        private const string ErrorMessageNoSelection = TextCollection.ResizeLabText.ErrorNoSelection;
        private const string ErrorMessageFewerThanTwoSelection = TextCollection.ResizeLabText.ErrorFewerThanTwoSelection;
        private const string ErrorMessageUndefined = TextCollection.ResizeLabText.ErrorUndefined;

        internal static bool IsSelecionValid(PowerPoint.Selection selection)
        {
            try
            {
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    ThrowErrorCode(ErrorCodeNoSelection);
                }

                return true;
            }
            catch (Exception e)
            {
                ProcessErrorMessage(e);
                return false;
            }
        }

        private static bool IsMoreThanOneShape(PowerPoint.ShapeRange selectedShapes)
        {
            try
            {
                if (selectedShapes.Count < 2)
                {
                    ThrowErrorCode(ErrorCodeFewerThanTwoSelection);
                }

                return true;
            }
            catch (Exception e)
            {
                ProcessErrorMessage(e);
                return false;
            }
        }

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
                case ErrorCodeFewerThanTwoSelection:
                    return ErrorMessageFewerThanTwoSelection;
                default:
                    return ErrorMessageUndefined;
            }
        }

        private static void IgnoreExceptionThrown() { }

        #endregion
    }
}
