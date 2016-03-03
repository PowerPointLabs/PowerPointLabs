using MahApps.Metro.Controls;
using System;
using System.Collections.Generic;
using System.Windows;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.PositionsLab
{
    /// <summary>
    /// Interaction logic for PositionsDistributeGridDialog.xaml
    /// </summary>
    public partial class PositionsDistributeGridDialog
    {
        //Flag to trigger
        public bool IsOpen { get; set; }

        //Private variables
        private int _numShapesSelected;
        private List<Shape> _selectedShapes;
        private int _rowLength;
        private int _colLength;

        public PositionsDistributeGridDialog(List<Shape> selectedShapes, int rowLength, int colLength)
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

        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            var rowValue = rowInput.Value;
            var colValue = colInput.Value;

            if (!rowValue.HasValue || rowValue.GetValueOrDefault() == 0 || 
                !colValue.HasValue || colValue.GetValueOrDefault() == 0)
            {
                return;
            }
            
            PositionsLabMain.DistributeGrid(_selectedShapes, (int)colValue, (int)rowValue);

            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
