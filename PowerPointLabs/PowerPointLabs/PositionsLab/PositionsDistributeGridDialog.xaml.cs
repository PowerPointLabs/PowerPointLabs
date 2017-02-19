using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.Utils;

namespace PowerPointLabs.PositionsLab
{
    /// <summary>
    /// Interaction logic for PositionsDistributeGridDialog.xaml
    /// </summary>
    public partial class PositionsDistributeGridDialog
    {
        //Error Messages
        private const string ErrorMessageNoSelection = TextCollection.PositionsLabText.ErrorNoSelection;
        private const string ErrorMessageFewerThanTwoSelection = TextCollection.PositionsLabText.ErrorFewerThanTwoSelection;
        private const string ErrorMessageFewerThanThreeSelection = TextCollection.PositionsLabText.ErrorFewerThanThreeSelection;
        private const string ErrorMessageFunctionNotSupportedForExtremeShapes = TextCollection.PositionsLabText.ErrorFunctionNotSupportedForWithinShapes;
        private const string ErrorMessageFunctionNotSupportedForSlide = TextCollection.PositionsLabText.ErrorFunctionNotSupportedForSlide;
        private const string ErrorMessageUndefined = TextCollection.PositionsLabText.ErrorUndefined;

        //Flag to trigger
        public bool IsOpen { get; set; }

        //Private variables
        private int _numShapesSelected;
        private ShapeRange _selectedShapes;
        private int _rowLength;
        private int _colLength;

        internal PositionsDistributeGridDialog(ShapeRange selectedShapes, int rowLength, int colLength)
        {
            IsOpen = true;
            _selectedShapes = selectedShapes;
            _numShapesSelected = selectedShapes.Count;
            _rowLength = rowLength;
            _colLength = colLength;
            InitializeComponent();
        }

        private void RowInput_Load(object sender, RoutedEventArgs e)
        {
            rowInput.Value = _colLength;
        }

        private void ColInput_Load(object sender, RoutedEventArgs e)
        {
            colInput.Value = _rowLength;
        }

        private void MarginTopInput_Load(object sender, RoutedEventArgs e)
        {
            marginTopInput.Value = PositionsLabMain.MarginTop;
        }

        private void MarginBottomInput_Load(object sender, RoutedEventArgs e)
        {
            marginBottomInput.Value = PositionsLabMain.MarginBottom;
        }

        private void MarginLeftInput_Load(object sender, RoutedEventArgs e)
        {
            marginLeftInput.Value = PositionsLabMain.MarginLeft;
        }

        private void MarginRightInput_Load(object sender, RoutedEventArgs e)
        {
            marginRightInput.Value = PositionsLabMain.MarginRight;
        }

        private void AlignLeftButton_Load(object sender, RoutedEventArgs e)
        {
            if (PositionsLabMain.DistributeGridAlignment == PositionsLabMain.GridAlignment.AlignLeft)
            {
                alignLeftButton.IsChecked = true;
            }
        }

        private void AlignCenterButton_Load(object sender, RoutedEventArgs e)
        {
            if (PositionsLabMain.DistributeGridAlignment == PositionsLabMain.GridAlignment.AlignCenter)
            {
                alignCenterButton.IsChecked = true;
            }
        }

        private void AlignRightButton_Load(object sender, RoutedEventArgs e)
        {
            if (PositionsLabMain.DistributeGridAlignment == PositionsLabMain.GridAlignment.AlignRight)
            {
                alignRightButton.IsChecked = true;
            }
        }

        private void RowInput_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double?> e)
        {
            if (colInput == null || rowInput == null)
            {
                return;
            }

            var value = rowInput.Value;

            if (!value.HasValue)
            {
                return;
            }

            var col = (int)Math.Ceiling(_numShapesSelected / value.GetValueOrDefault());
            colInput.Value = col;
        }

        private void ColInput_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double?> e)
        {
            if (rowInput == null || colInput == null)
            {
                return;
            }

            var value = colInput.Value;

            if (!value.HasValue)
            {
                return;
            }

            var row = (int)Math.Ceiling(_numShapesSelected / value.GetValueOrDefault());
            rowInput.Value = row;
        }

        private void PositionsDistributeGridDialong_Closed(object sender, EventArgs e)
        {
            IsOpen = false;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            var marginTopValue = marginTopInput.Value;
            var marginBottomValue = marginBottomInput.Value;
            var marginLeftValue = marginLeftInput.Value;
            var marginRightValue = marginRightInput.Value;

            // Checks for margin values
            if (!marginTopValue.HasValue || marginTopValue.GetValueOrDefault() < 0 ||
                !marginBottomValue.HasValue || marginBottomValue.GetValueOrDefault() < 0 ||
                !marginLeftValue.HasValue || marginLeftValue.GetValueOrDefault() < 0 ||
                !marginRightValue.HasValue || marginRightValue.GetValueOrDefault() < 0)
            {
                // TODO: Notify the user that not successfully changed
                return;
            }

            PositionsLabMain.SetDistributeMarginTop((float)marginTopValue);
            PositionsLabMain.SetDistributeMarginBottom((float)marginBottomValue);
            PositionsLabMain.SetDistributeMarginLeft((float)marginLeftValue);
            PositionsLabMain.SetDistributeMarginRight((float)marginRightValue);

            // Checks for distribute grid align reference
            if (alignLeftButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.SetDistributeGridAlignment(PositionsLabMain.GridAlignment.AlignLeft);
            }

            if (alignCenterButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.SetDistributeGridAlignment(PositionsLabMain.GridAlignment.AlignCenter);
            }

            if (alignRightButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.SetDistributeGridAlignment(PositionsLabMain.GridAlignment.AlignRight);
            }

            ShapeRange simulatedShapes = null;
            try
            {
                var rowValue = rowInput.Value;
                var colValue = colInput.Value;

                if (!rowValue.HasValue || rowValue.GetValueOrDefault() == 0 ||
                    !colValue.HasValue || colValue.GetValueOrDefault() == 0)
                {
                    return;
                }

                this.StartNewUndoEntry();
                var selectedShapes = this.GetCurrentSelection().ShapeRange;
                simulatedShapes = DuplicateShapes(selectedShapes);
                var simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);

                PositionsLabMain.DistributeGrid(simulatedPPShapes, (int)colValue, (int)rowValue);

                SyncShapes(selectedShapes, simulatedShapes);
            }
            catch (Exception ex)
            {
                Close();
                ShowErrorMessageBox(ex.Message, ex);
            }
            finally
            {
                simulatedShapes.Delete();
                GC.Collect();
            }

            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            IsOpen = false;
            Close();
        }

        private void SyncShapes(ShapeRange selected, ShapeRange simulatedShapes)
        {
            for (int i = 1; i <= selected.Count; i++)
            {
                var selectedShape = selected[i];
                var simulatedShape = simulatedShapes[i];
                var selectedCenter = Graphics.GetCenterPoint(selectedShape);
                var simulatedCenter = Graphics.GetCenterPoint(simulatedShape);

                selectedShape.IncrementLeft(simulatedCenter.X - selectedCenter.X);
                selectedShape.IncrementTop(simulatedCenter.Y - selectedCenter.Y);
            }
        }

        private ShapeRange DuplicateShapes(ShapeRange range)
        {
            String[] duplicatedShapeNames = new String[range.Count];

            for (int i = 0; i < range.Count; i++)
            {
                var shape = range[i + 1];
                var duplicated = shape.Duplicate()[1];
                duplicated.Name = shape.Name + "_Copy";
                duplicated.Left = shape.Left;
                duplicated.Top = shape.Top;
                duplicatedShapeNames[i] = duplicated.Name;
            }

            return this.GetCurrentSlide().Shapes.Range(duplicatedShapeNames);
        }

        private List<PPShape> ConvertShapeRangeToPPShapeList(ShapeRange range, int index)
        {
            var shapes = new List<PPShape>();

            for (var i = index; i <= range.Count; i++)
            {
                shapes.Add(new PPShape(range[i]));
            }

            return shapes;
        }
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
                Views.ErrorDialogWrapper.ShowDialog("Error", content, exception);
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

    }
}
