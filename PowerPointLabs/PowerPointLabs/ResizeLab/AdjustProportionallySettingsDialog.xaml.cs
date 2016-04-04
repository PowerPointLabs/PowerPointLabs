using System;
using System.Windows;

namespace PowerPointLabs.ResizeLab
{
    /// <summary>
    /// Interaction logic for AdjustProportionallySettingsDialog.xaml
    /// </summary>
    public partial class AdjustProportionallySettingsDialog
    {
        public bool IsOpen { get; set; }

        private readonly ResizeLabMain _resizeLab;
        public AdjustProportionallySettingsDialog(ResizeLabMain resizeLab)
        {
            _resizeLab = resizeLab;
            InitializeComponent();
            ResizeFactorTextBox.Text = _resizeLab.AdjustProportionallyResizeFactor.ToString("N2");
        }

        private void AdjustProportionallySettingsDialog_Closed(object sender, EventArgs e)
        {
            IsOpen = false;
        }

        private void OkBtn_Click(object sender, RoutedEventArgs e)
        {
            var resizeFactor = ResizeLabUtil.ConvertToFloat(ResizeFactorTextBox.Text);
            if (ResizeLabUtil.IsValidFactor(resizeFactor))
            {
                _resizeLab.AdjustProportionallyResizeFactor = (float) resizeFactor;
                Close();
            }
            else
            {
                MessageBox.Show("Please enter a value greater than 0", "Error");
            }
        }
    }

}
