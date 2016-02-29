using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PowerPointLabs.PositionsLab
{
    public partial class AlignSettingsForm : Form
    {
        public AlignSettingsForm()
        {
            InitializeComponent();
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            switch (alignReferenceSettings.SelectedIndex)
            {
                case 0: // Refer to slide
                    PositionsLabMain.AlignReferToSlide();
                    break;
                case 1: // Refer to shape
                    PositionsLabMain.AlignReferToShape();
                    break;
                default:
                    // Do nothing
                    // TODO: Throw error?
                    break;
            }
            this.Close();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
