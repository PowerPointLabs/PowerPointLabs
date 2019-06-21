using System;
using System.Windows;
using System.Windows.Input;
using PowerPointLabs.Utils.Windows;

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
            float? resizeFactor = ResizeLabUtil.ConvertToFloat(ResizeFactorTextBox.Text);
            if (ResizeLabUtil.IsValidFactor(resizeFactor))
            {
                _resizeLab.SlightAdjustResizeFactor = (float)resizeFactor;
                Close();
            }
            else
            {
                MessageBoxUtil.Show(TextCollection.ResizeLabText.ErrorValueLessThanEqualsZero, TextCollection.CommonText.ErrorTitle);
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
