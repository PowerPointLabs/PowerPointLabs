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
            UpdateYValues();
            UpdateLabels();
            UpdateColumnColors();
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

        private void UpdateColumnColors()
        {
            chart1.Series["Series1"].Points[0].Color = Color.FromArgb(255, 255, 0, 0);
            chart1.Series["Series1"].Points[1].Color = Color.FromArgb(255, 0, 255, 0);
            chart1.Series["Series1"].Points[2].Color = Color.FromArgb(255, 0, 0, 255);
        }

        private void UpdateLabels()
        {
            chart1.Series["Series1"].Points[0].Label = String.Format("{0:F}%", ((double)_selectedColor.R * 100 / 255.0f));
            chart1.Series["Series1"].Points[1].Label = String.Format("{0:F}%", ((double)_selectedColor.G * 100 / 255.0f));
            chart1.Series["Series1"].Points[2].Label = String.Format("{0:F}%", ((double)_selectedColor.B * 100 / 255.0f));
        }

        private void UpdateYValues()
        {
            chart1.Series["Series1"].Points[0].YValues[0] = (double)_selectedColor.R * 100 / 255.0f;
            chart1.Series["Series1"].Points[1].YValues[0] = (double)_selectedColor.G * 100 / 255.0f;
            chart1.Series["Series1"].Points[2].YValues[0] = (double)_selectedColor.B * 100 / 255.0f;
        } 
    }
}
