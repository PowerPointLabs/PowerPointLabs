using System;
using System.Collections.Generic;
using System.Windows;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.Utils;
using PowerPointLabs.Views;

namespace PowerPointLabs.PositionsLab.Views
{
    /// <summary>
    /// Interaction logic for PositionsDistributeGridDialog.xaml
    /// </summary>
    public partial class PositionsDistributeGridDialog
    {
        public delegate void DialogConfirmedDelegate(int rowLength, int colLength, 
                                                    PositionsLabSettings.GridAlignment gridAlignment,
                                                    float gridMarginTop, float gridMarginBottom,
                                                    float gridMarginLeft, float gridMarginRight);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }

        //Error Messages
        private const string ErrorMessageNoSelection = TextCollection.PositionsLabText.ErrorNoSelection;
        private const string ErrorMessageFewerThanTwoSelection = TextCollection.PositionsLabText.ErrorFewerThanTwoSelection;
        private const string ErrorMessageFewerThanThreeSelection = TextCollection.PositionsLabText.ErrorFewerThanThreeSelection;
        private const string ErrorMessageFunctionNotSupportedForExtremeShapes = TextCollection.PositionsLabText.ErrorFunctionNotSupportedForWithinShapes;
        private const string ErrorMessageFunctionNotSupportedForSlide = TextCollection.PositionsLabText.ErrorFunctionNotSupportedForSlide;
        private const string ErrorMessageUndefined = TextCollection.PositionsLabText.ErrorUndefined;

        //Private variables
        private int _numShapesSelected;

        #region Error Handling
        public void ShowErrorMessageBox(string content, Exception exception = null)
        {

            if (exception == null)
            {
                MessageBox.Show(content, "Error");
                return;
            }

            var errorMessage = GetErrorMessage(exception.Message);
            if (!string.Equals(errorMessage, ErrorMessageUndefined, StringComparison.Ordinal))
            {
                MessageBox.Show(content, "Error");
            }
            else
            {
                ErrorDialogBox.ShowDialog("Error", content, exception);
            }
        }

        private string GetErrorMessage(string errorMsg)
        {
            switch (errorMsg)
            {
                case ErrorMessageNoSelection:
                    return ErrorMessageNoSelection;
                case ErrorMessageFewerThanTwoSelection:
                    return ErrorMessageFewerThanTwoSelection;
                case ErrorMessageFewerThanThreeSelection:
                    return ErrorMessageFewerThanThreeSelection;
                case ErrorMessageFunctionNotSupportedForExtremeShapes:
                    return ErrorMessageFunctionNotSupportedForExtremeShapes;
                case ErrorMessageFunctionNotSupportedForSlide:
                    return ErrorMessageFunctionNotSupportedForSlide;
                default:
                    return ErrorMessageUndefined;
            }
        }

        private void IgnoreExceptionThrown() { }

        #endregion

        public PositionsDistributeGridDialog()
        {
            InitializeComponent();
        }

        public PositionsDistributeGridDialog(int numShapesSelected, 
                                            int rowLength, int colLength,
                                            PositionsLabSettings.GridAlignment gridAlignment,
                                            float gridMarginTop, float gridMarginBottom, 
                                            float gridMarginLeft, float gridMarginRight)
            : this()
        {
            _numShapesSelected = numShapesSelected;

            rowInput.Value = rowLength;
            colInput.Value = colLength;

            switch (gridAlignment)
            {
                case PositionsLabSettings.GridAlignment.AlignLeft:
                    alignLeftButton.IsChecked = true;
                    break;
                case PositionsLabSettings.GridAlignment.AlignCenter:
                    alignCenterButton.IsChecked = true;
                    break;
                case PositionsLabSettings.GridAlignment.AlignRight:
                    alignRightButton.IsChecked = true;
                    break;
            }

            marginTopInput.Value = gridMarginTop;
            marginBottomInput.Value = gridMarginBottom;
            marginLeftInput.Value = gridMarginLeft;
            marginRightInput.Value = gridMarginRight;
        }

        private void RowInput_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double?> e)
        {
            if (colInput == null || rowInput == null)
            {
                return;
            }

            double? value = rowInput.Value;

            if (!value.HasValue)
            {
                return;
            }

            int col = (int)Math.Ceiling(_numShapesSelected / value.GetValueOrDefault());
            colInput.Value = col;
        }

        private void ColInput_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double?> e)
        {
            if (colInput == null || rowInput == null)
            {
                return;
            }

            double? value = colInput.Value;

            if (!value.HasValue)
            {
                return;
            }

            int row = (int)Math.Ceiling(_numShapesSelected / value.GetValueOrDefault());
            rowInput.Value = row;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            double? marginTopValue = marginTopInput.Value;
            double? marginBottomValue = marginBottomInput.Value;
            double? marginLeftValue = marginLeftInput.Value;
            double? marginRightValue = marginRightInput.Value;

            // Checks for margin values
            if (!marginTopValue.HasValue || marginTopValue.GetValueOrDefault() < 0 ||
                !marginBottomValue.HasValue || marginBottomValue.GetValueOrDefault() < 0 ||
                !marginLeftValue.HasValue || marginLeftValue.GetValueOrDefault() < 0 ||
                !marginRightValue.HasValue || marginRightValue.GetValueOrDefault() < 0)
            {
                // TODO: Notify the user that not successfully changed
                return;
            }
            
            // Checks for distribute grid align reference
            PositionsLabSettings.GridAlignment gridAlignment;
            if (alignLeftButton.IsChecked.GetValueOrDefault())
            {
                gridAlignment = PositionsLabSettings.GridAlignment.AlignLeft;
            }
            else if (alignCenterButton.IsChecked.GetValueOrDefault())
            {
                gridAlignment = PositionsLabSettings.GridAlignment.AlignCenter;
            }
            else
            {
                gridAlignment = PositionsLabSettings.GridAlignment.AlignRight;
            }

            var rowValue = rowInput.Value;
            var colValue = colInput.Value;

            if (!rowValue.HasValue || rowValue.GetValueOrDefault() == 0 ||
                !colValue.HasValue || colValue.GetValueOrDefault() == 0)
            {
                // TODO: Notify the user that not successfully changed
                return;
            }

            DialogConfirmedHandler((int)rowValue, (int)colValue, gridAlignment, 
                                (float)marginTopValue, (float)marginBottomValue, 
                                (float)marginLeftValue, (float)marginRightValue);
            Close();
        }
    }
}
