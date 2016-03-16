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

        private ResizeLabMain _resizeLab;
        private Dictionary<ResizeLabMain.SameDimensionAnchor, RadioButton> anchorButtonLookUp;

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
            anchorButtonLookUp = new Dictionary<ResizeLabMain.SameDimensionAnchor, RadioButton>()
            {
                { ResizeLabMain.SameDimensionAnchor.TopLeft, AnchorTopLeftBtn},
                { ResizeLabMain.SameDimensionAnchor.Top, AnchorTopLeftBtn},
                { ResizeLabMain.SameDimensionAnchor.TopRight, AnchorTopRightBtn},
                { ResizeLabMain.SameDimensionAnchor.Left, AnchorLeftBtn},
                { ResizeLabMain.SameDimensionAnchor.Middle, AnchorMidBtn},
                { ResizeLabMain.SameDimensionAnchor.Right, AnchorRightBtn},
                { ResizeLabMain.SameDimensionAnchor.BottomLeft, AnchorBottomLeftBtn},
                { ResizeLabMain.SameDimensionAnchor.Bottom, AnchorBottomBtn},
                { ResizeLabMain.SameDimensionAnchor.BottomRight, AnchorBottomRightBtn}
            };
        }

        private void LoadAnchorCheckedButton()
        {
            RadioButton toCheckButton;
            if (anchorButtonLookUp.TryGetValue(_resizeLab.SameDimensionAnchorType, out toCheckButton))
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
            foreach (var anAnchorButton in anchorButtonLookUp)
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
