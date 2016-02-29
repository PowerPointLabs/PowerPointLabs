using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PowerPointLabs.PositionsLab
{
    public partial class DistributeSettingsForm : Form
    {
        public DistributeSettingsForm()
        {
            InitializeComponent();
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            switch (distributeReferenceSettings.SelectedIndex)
            {
                case 0: // Refer to slide
                    PositionsLabMain.DistributeReferToSlide();
                    break;
                case 1: // Refer to shape
                    PositionsLabMain.DistributeReferToShape();
                    break;
                default:
                    // Do nothing
                    // TODO: Throw exception?
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
