using System;
using System.Collections.Generic;
using System.Windows.Controls;
using MahApps.Metro.Controls;

namespace PowerPointLabs.ResizeLab
{
    /// <summary>
    /// Interaction logic for ResizeSameDimensionSettingsDialog.xaml
    /// </summary>
    public partial class SameDimensionSettingsDialog
    {
        public bool IsOpen { get; set; }

        private readonly ResizeLabMain _resizeLab;
        private Dictionary<ResizeLabMain.SameDimensionAnchor, RadioButton> _anchorButtonLookUp;

        public SameDimensionSettingsDialog(ResizeLabMain resizeLab)
        {
            IsOpen = true;
            _resizeLab = resizeLab;
            InitializeComponent();
            InitAnchorButtonDictionary();
            LoadAnchorCheckedButton();
        }

        private void InitAnchorButtonDictionary()
        {
            _anchorButtonLookUp = new Dictionary<ResizeLabMain.SameDimensionAnchor, RadioButton>()
            {
                { ResizeLabMain.SameDimensionAnchor.TopLeft, AnchorTopLeftBtn},
                { ResizeLabMain.SameDimensionAnchor.TopMiddle, AnchorTopMidBtn},
                { ResizeLabMain.SameDimensionAnchor.TopRight, AnchorTopRightBtn},
                { ResizeLabMain.SameDimensionAnchor.MiddleLeft, AnchorMidLeftBtn},
                { ResizeLabMain.SameDimensionAnchor.Middle, AnchorMidBtn},
                { ResizeLabMain.SameDimensionAnchor.MiddleRight, AnchorMidRightBtn},
                { ResizeLabMain.SameDimensionAnchor.BottomLeft, AnchorBottomLeftBtn},
                { ResizeLabMain.SameDimensionAnchor.BottomMiddle, AnchorBottomMidBtn},
                { ResizeLabMain.SameDimensionAnchor.BottomRight, AnchorBottomRightBtn}
            };
        }

        private void LoadAnchorCheckedButton()
        {
            RadioButton toCheckButton;
            if (_anchorButtonLookUp.TryGetValue(_resizeLab.SameDimensionAnchorType, out toCheckButton))
            {
                toCheckButton.IsChecked = true;
            }
        }

        private void SameDimensionSettingsDialog_Closed(object sender, EventArgs e)
        {
            IsOpen = false;
        }

        private void OkBtn_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            _resizeLab.SameDimensionAnchorType = CheckedAnchorBtnToAnchorType();

            IsOpen = false;
            Close();
        }

        #region Helper Functions

        private ResizeLabMain.SameDimensionAnchor CheckedAnchorBtnToAnchorType()
        {
            foreach (var anAnchorButton in _anchorButtonLookUp)
            {
                if (anAnchorButton.Value.IsChecked.GetValueOrDefault())
                {
                    return anAnchorButton.Key;
                }
                    
            }
            return ResizeLabMain.SameDimensionAnchor.Middle; // Should not execute
        }

        #endregion

    }
}
