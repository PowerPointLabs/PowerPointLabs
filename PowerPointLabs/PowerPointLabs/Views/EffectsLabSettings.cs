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
    public partial class EffectsLabSettings : Form
    {
        public delegate void UpdateSettingsDelegate(bool isCover);
        public UpdateSettingsDelegate SettingsHandler;

        public EffectsLabSettings(bool isCoverCheckInitial)
        {
            InitializeComponent();

            isCoverChecker.Checked = isCoverCheckInitial;
        }

        private void OkButtonClick(object sender, EventArgs e)
        {
            SettingsHandler(isCoverChecker.Checked);
            Close();
        }
    }
}
