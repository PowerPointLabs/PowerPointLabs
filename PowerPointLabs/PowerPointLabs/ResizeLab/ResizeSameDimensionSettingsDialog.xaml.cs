using System;
using System.Collections.Generic;
using System.Windows.Controls;

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
                { ResizeLabMain.SameDimensionAnchor.TopCenter, AnchorTopMidBtn},
                { ResizeLabMain.SameDimensionAnchor.TopRight, AnchorTopRightBtn},
                { ResizeLabMain.SameDimensionAnchor.MiddleLeft, AnchorMidLeftBtn},
                { ResizeLabMain.SameDimensionAnchor.Center, AnchorMidBtn},
                { ResizeLabMain.SameDimensionAnchor.MiddleRight, AnchorMidRightBtn},
                { ResizeLabMain.SameDimensionAnchor.BottomLeft, AnchorBottomLeftBtn},
                { ResizeLabMain.SameDimensionAnchor.BottomCenter, AnchorBottomMidBtn},
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
            _resizeLab.SameDimensionAnchorType = AnchorTypeToCheckedAnchorBtn();

            Close();
        }

        #region Helper Functions

        private ResizeLabMain.SameDimensionAnchor AnchorTypeToCheckedAnchorBtn()
        {
            foreach (var anAnchorButton in _anchorButtonLookUp)
            {
                if (anAnchorButton.Value.IsChecked.GetValueOrDefault())
                {
                    return anAnchorButton.Key;
                }
                    
            }
            return ResizeLabMain.SameDimensionAnchor.Center; // Should not execute
        }

        #endregion

    }
}
