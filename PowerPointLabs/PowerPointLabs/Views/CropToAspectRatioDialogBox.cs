using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PowerPointLabs.Views
{
    public partial class CropToAspectRatioDialogBox : Form
    {
#pragma warning disable 0618

        public CropToAspectRatioDialogBox()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            string widthText = widthTextBox.Text;
            string heightText = heightTextBox.Text;

            Globals.ThisAddIn.Ribbon.CropToAspectRatioInput(widthText, heightText);
            this.Close();
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
