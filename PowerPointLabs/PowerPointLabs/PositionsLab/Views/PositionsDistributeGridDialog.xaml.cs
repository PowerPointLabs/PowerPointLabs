using System;
using System.Windows;

using PowerPointLabs.TextCollection;
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

        //Private variables
        private int _numShapesSelected;

        #region Error Handling
        public void ShowErrorMessageBox(string content, Exception exception = null)
        {

            if (exception == null)
            {
                MessageBox.Show(content, PositionsLabText.ErrorDialogTitle);
                return;
            }

            string errorMessage = GetErrorMessage(exception.Message);
            if (!string.Equals(errorMessage, PositionsLabText.ErrorUndefined, StringComparison.Ordinal))
            {
                MessageBox.Show(content, PositionsLabText.ErrorDialogTitle);
            }
            else
            {
                ErrorDialogBox.ShowDialog(PositionsLabText.ErrorDialogTitle, content, exception);
            }
        }

        private string GetErrorMessage(string errorMsg)
        {
            switch (errorMsg)
            {
                case PositionsLabText.ErrorNoSelection:
                    return PositionsLabText.ErrorNoSelection;
                case PositionsLabText.ErrorFewerThanTwoSelection:
                    return PositionsLabText.ErrorFewerThanTwoSelection;
                case PositionsLabText.ErrorFewerThanThreeSelection:
                    return PositionsLabText.ErrorFewerThanThreeSelection;
                case PositionsLabText.ErrorFunctionNotSupportedForWithinShapes:
                    return PositionsLabText.ErrorFunctionNotSupportedForWithinShapes;
                case PositionsLabText.ErrorFunctionNotSupportedForSlide:
                    return PositionsLabText.ErrorFunctionNotSupportedForSlide;
                default:
                    return PositionsLabText.ErrorUndefined;
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
                ShowErrorMessageBox(PositionsLabText.ErrorRepositionFail);
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

            double? rowValue = rowInput.Value;
            double? colValue = colInput.Value;

            if (!rowValue.HasValue || rowValue.GetValueOrDefault() == 0 ||
                !colValue.HasValue || colValue.GetValueOrDefault() == 0)
            {
                ShowErrorMessageBox(PositionsLabText.ErrorRepositionFail);
                return;
            }

            DialogConfirmedHandler((int)rowValue, (int)colValue, gridAlignment, 
                                (float)marginTopValue, (float)marginBottomValue, 
                                (float)marginLeftValue, (float)marginRightValue);
            Close();
        }
    }
}
