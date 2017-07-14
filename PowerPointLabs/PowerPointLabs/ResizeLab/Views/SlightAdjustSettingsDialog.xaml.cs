using System;
using System.Windows;
using System.Windows.Input;

namespace PowerPointLabs.ResizeLab.Views
{
    /// <summary>
    /// Interaction logic for SlightAdjustSettingsDialog.xaml
    /// </summary>
    public partial class SlightAdjustSettingsDialog
    {
        public bool IsOpen { get; set; }

        private readonly ResizeLabMain _resizeLab;
        public SlightAdjustSettingsDialog(ResizeLabMain resizeLab)
        {
            _resizeLab = resizeLab;
            InitializeComponent();
            ResizeFactorTextBox.Text = _resizeLab.SlightAdjustResizeFactor.ToString("N2");
        }

        private void SlightAdjustSettingsDialog_Closed(object sender, EventArgs e)
        {
            IsOpen = false;
        }

        private void OkBtn_Click(object sender, RoutedEventArgs e)
        {
            var resizeFactor = ResizeLabUtil.ConvertToFloat(ResizeFactorTextBox.Text);
            if (ResizeLabUtil.IsValidFactor(resizeFactor))
            {
                _resizeLab.SlightAdjustResizeFactor = (float)resizeFactor;
                Close();
            }
            else
            {
                MessageBox.Show("Please enter a value greater than 0", "Error");
            }
        }

        private void Dialog_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Close();
            }
        }
    }

}
