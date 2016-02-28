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
    public partial class AlignSettingsForm : Form
    {
        public AlignSettingsForm()
        {
            InitializeComponent();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void AlignReferenceSettings_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
