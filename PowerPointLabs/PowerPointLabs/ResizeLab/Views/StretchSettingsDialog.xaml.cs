using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace PowerPointLabs.ResizeLab.Views
{
    /// <summary>
    /// Interaction logic for StretchSettingsDialog.xaml
    /// </summary>
    public partial class StretchSettingsDialog
    {
        public bool IsOpen { get; set; }

        private readonly ResizeLabMain _resizeLab;
        private Dictionary<ResizeLabMain.StretchRefType, RadioButton> _refTypeButtonLookUp;

        public StretchSettingsDialog(ResizeLabMain resizeLab)
        {
            IsOpen = true;
            _resizeLab = resizeLab;
            InitializeComponent();
            InitRefTypeButtonDictionary();
            LoadRefTypeCheckedButton();
        }

        #region Initialize
        private void InitRefTypeButtonDictionary()
        {
            _refTypeButtonLookUp = new Dictionary<ResizeLabMain.StretchRefType, RadioButton>()
            {
                { ResizeLabMain.StretchRefType.FirstSelected, FirstSelectedBtn },
                { ResizeLabMain.StretchRefType.Outermost, OuterMostShapeBtn }
            };
        }

        private void LoadRefTypeCheckedButton()
        {
            RadioButton toCheckButton;
            if (_refTypeButtonLookUp.TryGetValue(_resizeLab.ReferenceType, out toCheckButton))
            {
                toCheckButton.IsChecked = true;
            }
        }

        #endregion

        #region Event Handler
        private void OkBtn_Click(object sender, RoutedEventArgs e)
        {
            _resizeLab.ReferenceType = RefTypeToCheckedRefTypeBtn();
            
            Close();
        }

        private void StretchSettingsDialog_Closed(object sender, EventArgs e)
        {
            IsOpen = false;
        }

        private void Dialog_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Close();
            }
        }

        #endregion

        #region Helper functions
        private ResizeLabMain.StretchRefType RefTypeToCheckedRefTypeBtn()
        {
            foreach (var aRefTypeButton in _refTypeButtonLookUp)
            {
                if (aRefTypeButton.Value.IsChecked.GetValueOrDefault())
                {
                    return aRefTypeButton.Key;
                }
            }
            return ResizeLabMain.StretchRefType.FirstSelected; // Should not execute
        }

        #endregion

    }
}
