using PowerPointLabs.Utils;
using System;
using System.Collections.Generic;
using System.Windows;
using PowerPointLabs.ActionFramework.Common.Extension;
using Microsoft.Office.Interop.PowerPoint;

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

            this.StartNewUndoEntry();
            var selectedShapes = this.GetCurrentSelection().ShapeRange;
            var simulatedShapes = DuplicateShapes(selectedShapes);
            var simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);

            PositionsLabMain.DistributeGrid(simulatedPPShapes, (int)colValue, (int)rowValue);

            SyncShapes(selectedShapes, simulatedShapes);
            simulatedShapes.Delete();
            GC.Collect();

            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
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
    }
}
