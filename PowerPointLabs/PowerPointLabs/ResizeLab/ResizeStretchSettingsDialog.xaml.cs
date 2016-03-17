using System;
using System.Windows;

namespace PowerPointLabs.ResizeLab
{
    /// <summary>
    /// Interaction logic for StretchSettingsDialog.xaml
    /// </summary>
    public partial class StretchSettingsDialog
    {
        //Flag to trigger
        public bool IsOpen { get; set; }

        private ResizeLabMain _resizeLab;

        public StretchSettingsDialog(ResizeLabMain resizeLab)
        {
            IsOpen = true;
            _resizeLab = resizeLab;
            InitializeComponent();
        }

        private void OkBtn_Click(object sender, RoutedEventArgs e)
        {
            if (FirstSelectedBtn.IsChecked.HasValue && FirstSelectedBtn.IsChecked.Value)
            {
                _resizeLab.ReferenceType = ResizeLabMain.RefType.FirstSelected;
            }
            else if (OuterMostShapeBtn.IsChecked.HasValue && OuterMostShapeBtn.IsChecked.Value)
            {
                _resizeLab.ReferenceType = ResizeLabMain.RefType.Outermost;
            }
            IsOpen = false;
            Close();
        }

        private void StretchSettingsDialog_Closed(object sender, EventArgs e)
        {
            IsOpen = false;
        }

        private void FirstSelectedBtn_Load(object sender, RoutedEventArgs e)
        {
            if (_resizeLab != null && _resizeLab.ReferenceType == ResizeLabMain.RefType.FirstSelected)
            {
                FirstSelectedBtn.IsChecked = true;
            }
        }

        private void OuterMostShapeBtn_Load(object sender, RoutedEventArgs e)
        {
            if (_resizeLab != null && _resizeLab.ReferenceType == ResizeLabMain.RefType.Outermost)
            {
                OuterMostShapeBtn.IsChecked = true;
            }
        }
    }
}
