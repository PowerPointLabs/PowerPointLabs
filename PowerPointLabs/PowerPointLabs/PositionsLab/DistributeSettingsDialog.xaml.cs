using MahApps.Metro.Controls;
using Microsoft.Office.Interop.PowerPoint;
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

namespace PowerPointLabs.PositionsLab
{
    /// <summary>
    /// Interaction logic for DistributeSettingsDialog.xaml
    /// </summary>
    public partial class DistributeSettingsDialog : MetroWindow
    {
        // User Control
        private NumericUpDown _marginTopInput;
        private NumericUpDown _marginBottomInput;
        private NumericUpDown _marginLeftInput;
        private NumericUpDown _marginRightInput;
        private RadioButton _alignLeftButton;
        private RadioButton _alignCenterButton;
        private RadioButton _alignRightButton;


        public DistributeSettingsDialog()
        {
            InitializeComponent();
        }

        #region On-Load Settings
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
        #endregion

        #region Button actions
        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            var marginTopValue = _marginTopInput.Value;
            var marginBottomValue = _marginBottomInput.Value;
            var marginLeftValue = _marginLeftInput.Value;
            var marginRightValue = _marginRightInput.Value;

            if (distributeToShapeButton.IsChecked == true)
            {
                PositionsLabMain.DistributeReferToShape();
            }
            else
            {
                PositionsLabMain.DistributeReferToSlide();
            }
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

            if (_alignLeftButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.SetDistributeGridAlignment(PositionsLabMain.ALIGN_LEFT);
            }

            if (_alignCenterButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.SetDistributeGridAlignment(PositionsLabMain.ALIGN_CENTER);
            }

            if (_alignRightButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.SetDistributeGridAlignment(PositionsLabMain.ALIGN_RIGHT);
            }

            this.Hide();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
        }

        private void DistributeToShapeButton_Click(object sender, RoutedEventArgs e)
        {
            distributeToSlideButton.IsChecked = false;
        }

        private void DistributeToSlideButton_Click(object sender, RoutedEventArgs e)
        {
            distributeToShapeButton.IsChecked = false;
        }
        #endregion
    }
}
