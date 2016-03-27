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
        private Dictionary<ResizeLabMain.AnchorPoint, RadioButton> _anchorButtonLookUp;

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
            _anchorButtonLookUp = new Dictionary<ResizeLabMain.AnchorPoint, RadioButton>()
            {
                { ResizeLabMain.AnchorPoint.TopLeft, AnchorTopLeftBtn},
                { ResizeLabMain.AnchorPoint.TopCenter, AnchorTopMidBtn},
                { ResizeLabMain.AnchorPoint.TopRight, AnchorTopRightBtn},
                { ResizeLabMain.AnchorPoint.MiddleLeft, AnchorMidLeftBtn},
                { ResizeLabMain.AnchorPoint.Center, AnchorMidBtn},
                { ResizeLabMain.AnchorPoint.MiddleRight, AnchorMidRightBtn},
                { ResizeLabMain.AnchorPoint.BottomLeft, AnchorBottomLeftBtn},
                { ResizeLabMain.AnchorPoint.BottomCenter, AnchorBottomMidBtn},
                { ResizeLabMain.AnchorPoint.BottomRight, AnchorBottomRightBtn}
            };
        }

        private void LoadAnchorCheckedButton()
        {
            RadioButton toCheckButton;
            if (_anchorButtonLookUp.TryGetValue(_resizeLab.AnchorPointType, out toCheckButton))
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
            _resizeLab.AnchorPointType = AnchorTypeToCheckedAnchorBtn();

            Close();
        }

        #region Helper Functions

        private ResizeLabMain.AnchorPoint AnchorTypeToCheckedAnchorBtn()
        {
            foreach (var anAnchorButton in _anchorButtonLookUp)
            {
                if (anAnchorButton.Value.IsChecked.GetValueOrDefault())
                {
                    return anAnchorButton.Key;
                }
                    
            }
            return ResizeLabMain.AnchorPoint.Center; // Should not execute
        }

        #endregion

    }
}
