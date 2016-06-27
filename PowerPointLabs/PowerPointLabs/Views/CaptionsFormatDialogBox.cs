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
    public partial class CaptionsFormatDialogBox : Form
    {
        public delegate void UpdateSettingsDelegate(String a, bool preview);
        public UpdateSettingsDelegate SettingsHandler;

        public CaptionsFormatDialogBox()
        {
            InitializeComponent();
            this.ShowInTaskbar = false;
        }

        public CaptionsFormatDialogBox(int a, bool preview) : this()
        {

        }

        private void CaptionsFormatDialogBox_Load(object sender, EventArgs e)
        {
            
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Ok_Click(object sender, EventArgs e)
        {
            //SettingsHandler(1, true);
            Close();
        }
    }
}
