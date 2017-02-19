using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using PPExtraEventHelper;
using TestInterface;

namespace PowerPointLabs.ColorPicker
{
    public partial class ColorInformationDialog : Form, IColorsLabMoreInfoDialog
    {
        private HSLColor _selectedColor;
        
        public ColorInformationDialog(HSLColor selectedColor)
        {
            _selectedColor = selectedColor;
            InitializeComponent();
            SetUpUI();
        }

        private void TextBox_Enter(object sender, EventArgs e)
        {
            var textBox = (TextBox) sender;
            Clipboard.SetText(textBox.Text.Substring(5));
            if (textBox.Equals(hslTextBox))
            {
                label1.Text = "HSL value copied";
            } 
            else if (textBox.Equals(rgbTextBox))
            {
                label1.Text = "RGB value copied";
            } 
            else if (textBox.Equals(hexTextBox))
            {
                label1.Text = "HEX value copied";
            }
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
            toolTip1.SetToolTip(this.hslTextBox, "Hue, Saturation, Luminance");
            toolTip1.SetToolTip(this.hexTextBox, "Hex Triplet");
        }

        private void UpdateHSLTextBox()
        {
            hslTextBox.Text = String.Format("HSL: {0}, {1}, {2}", (int)_selectedColor.Hue,
            (int) (_selectedColor.Saturation), (int) (_selectedColor.Luminosity));
            hslTextBox.Enter += TextBox_Enter;
        }

        private void UpdateRGBTextBox()
        {
            rgbTextBox.Text = String.Format("RGB: {0}, {1}, {2}", ((Color)_selectedColor).R,
            ((Color)_selectedColor).G, ((Color)_selectedColor).B);
            rgbTextBox.Enter += TextBox_Enter;
        }

        private void UpdateHexTextBox()
        {
            byte[] rgbArray = { ((Color)_selectedColor).R, ((Color)_selectedColor).G, ((Color)_selectedColor).B };
            hexTextBox.Text = "HEX: #" + ByteArrayToString(rgbArray);
            hexTextBox.Enter += TextBox_Enter;
        }

        private string ByteArrayToString(byte[] ba)
        {
            string hex = BitConverter.ToString(ba);
            return hex.Replace("-", "");
        }

        #region Functional Test method
        public string GetHslText()
        {
            return hslTextBox.Text;
        }

        public string GetRgbText()
        {
            return rgbTextBox.Text;
        }

        public string GetHexText()
        {
            return hexTextBox.Text;
        }

        public void TearDown()
        {
            Close();
        }
        # endregion
    }
}
