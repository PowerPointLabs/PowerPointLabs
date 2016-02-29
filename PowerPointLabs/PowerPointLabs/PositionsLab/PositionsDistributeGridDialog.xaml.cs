using MahApps.Metro.Controls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using System.Diagnostics;

namespace PowerPointLabs.PositionsLab
{
    /// <summary>
    /// Interaction logic for PositionsDistributeGridDialog.xaml
    /// </summary>
    public partial class PositionsDistributeGridDialog : MetroWindow
    {
        //Flag to trigger
        public bool IsOpen { get; set; }

        //User Control
        private NumericUpDown _rowInput;
        private NumericUpDown _colInput;
        private NumericUpDown _marginTopInput;
        private NumericUpDown _marginBottomInput;
        private NumericUpDown _marginLeftInput;
        private NumericUpDown _marginRightInput;
        private RadioButton _alignLeftButton;
        private RadioButton _alignCenterButton;
        private RadioButton _alignRightButton;

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
            _rowInput = (NumericUpDown)sender;
            _rowInput.Value = _rowLength;
        }

        private void ColInput_Load(object sender, RoutedEventArgs e)
        {
            _colInput = (NumericUpDown)sender;
            _colInput.Value = _colLength;
        }

        private void MarginTopInput_Load(object sender, RoutedEventArgs e)
        {
            _marginTopInput = (NumericUpDown)sender;
        }

        private void MarginBottomInput_Load(object sender, RoutedEventArgs e)
        {
            _marginBottomInput = (NumericUpDown)sender;
        }

        private void MarginLeftInput_Load(object sender, RoutedEventArgs e)
        {
            _marginLeftInput = (NumericUpDown)sender;
        }

        private void MarginRightInput_Load(object sender, RoutedEventArgs e)
        {
            _marginRightInput = (NumericUpDown)sender;
        }

        private void AlignLeftButton_Load(object sender, RoutedEventArgs e)
        {
            _alignLeftButton = (RadioButton)sender;
        }

        private void AlignCenterButton_Load(object sender, RoutedEventArgs e)
        {
            _alignCenterButton = (RadioButton)sender;
        }

        private void AlignRightButton_Load(object sender, RoutedEventArgs e)
        {
            _alignRightButton = (RadioButton)sender;
        }

        private void RowInput_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double?> e)
        {
            if (_colInput == null || _rowInput == null)
            {
                return;
            }

            var value = _rowInput.Value;

            if (!value.HasValue)
            {
                return;
            }

            int col = (int)Math.Ceiling(_numShapesSelected / value.GetValueOrDefault());
            _colInput.Value = col;
        }

        private void ColInput_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double?> e)
        {
            if (_colInput == null || _rowInput == null)
            {
                return;
            }

            var value = _colInput.Value;

            if (!value.HasValue)
            {
                return;
            }

            int row = (int)Math.Ceiling(_numShapesSelected / value.GetValueOrDefault());
            _rowInput.Value = row;
        }

        private void PositionsDistributeGridDialong_Closed(object sender, EventArgs e)
        {
            IsOpen = false;
        }

        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            var rowValue = _rowInput.Value;
            var colValue = _colInput.Value;
            var marginTopValue = _marginTopInput.Value;
            var marginBottomValue = _marginBottomInput.Value;
            var marginLeftValue = _marginLeftInput.Value;
            var marginRightValue = _marginRightInput.Value;
            int alignment = -1;

            if (!rowValue.HasValue || rowValue.GetValueOrDefault() == 0 || 
                !colValue.HasValue || colValue.GetValueOrDefault() == 0)
            {
                return;
            }

            if (!marginTopValue.HasValue || marginTopValue.GetValueOrDefault() < 0 || 
                !marginBottomValue.HasValue || marginBottomValue.GetValueOrDefault() < 0 ||
                !marginLeftValue.HasValue || marginLeftValue.GetValueOrDefault() < 0 ||
                !marginRightValue.HasValue || marginRightValue.GetValueOrDefault() < 0)
            {
                return;
            }

            if (_alignLeftButton.IsChecked.GetValueOrDefault())
            {
                alignment = PositionsLabMain.ALIGN_LEFT;
            }

            if (_alignCenterButton.IsChecked.GetValueOrDefault())
            {
                alignment = PositionsLabMain.ALIGN_CENTER;
            }

            if (_alignRightButton.IsChecked.GetValueOrDefault())
            {
                alignment = PositionsLabMain.ALIGN_RIGHT;
            }

            if (alignment < 0)
            {
                return;
            }

            PositionsLabMain.DistributeGrid(_selectedShapes, (int)rowValue, (int)colValue, (float)marginTopValue, 
                (float)marginBottomValue, (float)marginLeftValue, (float)marginRightValue, alignment);

            this.Close();
        }
    }
}
