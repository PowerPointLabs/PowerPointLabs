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

    public partial class AutoCaptionDialogBox : Form
    {
        public delegate void UpdateSettingsDelegate(bool allSlides);
        public UpdateSettingsDelegate SettingsHandler;

        public AutoCaptionDialogBox()
        {
            InitializeComponent();
            this.ShowInTaskbar = false;
        }

        public AutoCaptionDialogBox(bool allSlides) : this()
        {
            this.allSlides.Checked = allSlides;
        }

        private void cancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void ok_Click(object sender, EventArgs e)
        {
            SettingsHandler(allSlides.Checked);
            Close();
        }
    }
}
