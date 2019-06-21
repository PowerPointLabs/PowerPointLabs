using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using PowerPointLabs.Utils.Windows;

namespace PowerPointLabs.ResizeLab.Views
{
    /// <summary>
    /// Interaction logic for AdjustProportionallySettingsDialog.xaml
    /// </summary>
    public partial class AdjustProportionallySettingsDialog
    {
        public bool IsOpen { get; set; }

        private const double RowHeight = 30;
        private const int ShapeGridColumnIndex = 2;

        private readonly ResizeLabMain _resizeLab;
        public AdjustProportionallySettingsDialog(ResizeLabMain resizeLab, int noOfShapes)
        {
            if (noOfShapes < 2)
            {
                return;
            }
            _resizeLab = resizeLab;
            InitializeComponent();
            AddShapeRows(noOfShapes);
        }

        #region Initialise
        private void AddShapeRows(int noOfShapes)
        {
            if (noOfShapes >= 3)
            {
                AddShapeRow("3rd Selected Object");
            }
            for (int i = 4; i <= noOfShapes; i++)
            {
                AddShapeRow(i + "th Selected Object");
            }
        }

        private void AddShapeRow(string label)
        {
            // Increase height of main grid
            double oldHeight = MainGrid.RowDefinitions[ShapeGridColumnIndex].Height.Value;
            MainGrid.RowDefinitions[ShapeGridColumnIndex].Height = new GridLength(oldHeight + RowHeight);
            Height += RowHeight;

            // Add a row to inner grid
            RowDefinition newShapeRow = new RowDefinition();
            newShapeRow.Height = new GridLength(1, GridUnitType.Star);
            ShapesGrid.RowDefinitions.Add(newShapeRow);

            // Prepare the element
            TextBlock labelTextBlock = new TextBlock();
            labelTextBlock.Text = label;
            labelTextBlock.VerticalAlignment = VerticalAlignment.Center;
            labelTextBlock.HorizontalAlignment = HorizontalAlignment.Right;

            TextBox textBox = new TextBox();
            textBox.Width = 50;
            textBox.VerticalAlignment = VerticalAlignment.Center;
            textBox.HorizontalAlignment = HorizontalAlignment.Center;
            textBox.ToolTip = ResizeLabTooltip.AdjustProportionallySettingsTextBox;

            // Append the element
            int rowIndex = ShapesGrid.RowDefinitions.Count - 1;
            ShapesGrid.Children.Add(labelTextBlock);
            ShapesGrid.Children.Add(textBox);

            Grid.SetColumn(labelTextBlock, 0);
            Grid.SetRow(labelTextBlock, rowIndex);
            Grid.SetColumn(textBox, 1);
            Grid.SetRow(textBox, rowIndex);
        }

        #endregion

        #region Event Handler
        private void AdjustProportionallySettingsDialog_Closed(object sender, EventArgs e)
        {
            IsOpen = false;
        }

        private void OkBtn_Click(object sender, RoutedEventArgs e)
        {
            List<float> proportionList = new List<float>();
            for (int i = 1; i < ShapesGrid.Children.Count; i += 2)
            {
                TextBox textBox = ShapesGrid.Children[i] as TextBox;
                float? proportion = ResizeLabUtil.ConvertToFloat(textBox.Text);

                if (ResizeLabUtil.IsValidFactor(proportion))
                {
                    proportionList.Add(proportion.Value);
                }
                else
                {
                    MessageBoxUtil.Show(string.Format(TextCollection.ResizeLabText.ErrorValueLessThanEqualsZeroWithShape, (i + 1)/2), TextCollection.CommonText.ErrorTitle);
                    return;
                }
            }
            _resizeLab.AdjustProportionallyProportionList = proportionList;
            Close();
        }

        private void Dialog_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Close();
            }
        }

        #endregion
    }

}
