using System;
using System.Globalization;
using System.Linq;
using System.Windows;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    internal partial class ResizeLabMain
    {
        private const string MessageBoxTitle = "Error during Resizing";

        private const int ErrorCodeNoSelection = 0;
        private const int ErrorCodeFewerThanTwoSelection = 1;
        private const int ErrorCodeShapesNotStretchText = 2;

        private const string ErrorMessageNoSelection = TextCollection.ResizeLabText.ErrorNoSelection;
        private const string ErrorMessageFewerThanTwoSelection = TextCollection.ResizeLabText.ErrorFewerThanTwoSelection;
        private const string ErrorMessageUndefined = TextCollection.ResizeLabText.ErrorUndefined;
        private const string ErrorMessageShapesNotStretchText =
            TextCollection.ResizeLabText.WarningShapesNotStretchText;

        private IResizeLabPane View { get; }

        public ResizeLabMain(IResizeLabPane view)
        {
            View = view;
        }

        private enum Dimension
        {
            Height,
            Width,
            HeightAndWidth
        }

        internal bool IsShapeSelection(PowerPoint.Selection selection)
        {
            return selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes;
        }

        internal bool IsSelecionValid(PowerPoint.Selection selection)
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

        private bool IsMoreThanOneShape(PowerPoint.ShapeRange selectedShapes)
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

        private void ThrowErrorCode(int errorType)
        {
            throw new Exception(errorType.ToString(CultureInfo.InvariantCulture));
        }

        private void ProcessErrorMessage(Exception e)
        {
            var errorMessage = GetErrorMessage(e.Message);
            if (!string.Equals(errorMessage, ErrorMessageUndefined, StringComparison.Ordinal))
            {
                View.ShowErrorMessageBox(errorMessage);
            }
            else
            {
                View.ShowErrorMessageBox(e.Message, e);
            }
        }

        private string GetErrorMessage(string errorCode)
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
                case ErrorCodeShapesNotStretchText:
                    return ErrorMessageShapesNotStretchText;
                default:
                    return ErrorMessageUndefined;
            }
        }

        private void IgnoreExceptionThrown() { }

        #endregion
    }
}
