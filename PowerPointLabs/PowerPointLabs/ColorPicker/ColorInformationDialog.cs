using PPExtraEventHelper;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PowerPointLabs.ColorPicker
{
    public partial class ColorInformationDialog : Form
    {
        private Color _selectedColor;
        
        public ColorInformationDialog(Color selectedColor)
        {
            _selectedColor = selectedColor;
            InitializeComponent();
            SetUpUI();
        }

        private void textBox_GotFocus(object sender, EventArgs e)
        {
            Native.HideCaret(((TextBox)sender).Handle);
        }

        private void SetUpUI()
        {
            selectedColorPanel.BackColor = _selectedColor;
            UpdateHexTextBox();
            UpdateRGBTextBox();
            UpdateHSLTextBox();
            SetUpToolTips();
        }

        private void SetUpToolTips()
        {
            toolTip1.SetToolTip(this.rgbTextBox, "Red, Blue, Green");
            toolTip1.SetToolTip(this.HSLTextBox, "Hue, Saturation, Luminance");
            toolTip1.SetToolTip(this.hexTextBox, "Hex Triplet");
        }

        private void UpdateHSLTextBox()
        {
            HSLTextBox.Text = String.Format("HSL ({0:F}" + ((char)176) + ", {1:F}, {2:F})", _selectedColor.GetHue(),
            _selectedColor.GetSaturation(), _selectedColor.GetBrightness());
            HSLTextBox.GotFocus += textBox_GotFocus;
        }

        private void UpdateRGBTextBox()
        {
            rgbTextBox.Text = String.Format("RGB ({0}, {1}, {2})", _selectedColor.R,
            _selectedColor.G, _selectedColor.B);
            rgbTextBox.GotFocus += textBox_GotFocus;
        }

        private void UpdateHexTextBox()
        {
            byte[] rgbArray = { _selectedColor.R, _selectedColor.G, _selectedColor.B };
            hexTextBox.Text = "#" + ByteArrayToString(rgbArray);
            hexTextBox.GotFocus += textBox_GotFocus;
        }

        private string ByteArrayToString(byte[] ba)
        {
            string hex = BitConverter.ToString(ba);
            return hex.Replace("-", "");
        }
    }
}
