using System;
using System.Windows.Forms;

namespace PowerPointLabs.Views
{
    public partial class AutoZoomDialogBox : Form
    {
        public delegate void UpdateSettingsDelegate(bool slideBackgroundChecked, bool multiSlideChecked);
        public UpdateSettingsDelegate SettingsHandler;
        public AutoZoomDialogBox()
        {
            InitializeComponent();
        }

        public AutoZoomDialogBox(bool backgroundChecked, bool multiChecked)
            : this()
        {
            this.checkBox1.Checked = backgroundChecked;
            this.checkBox2.Checked = multiChecked;
        }

        private void AutoZoomDialogBox_Load(object sender, EventArgs e)
        {
            ToolTip ttCheckBox = new ToolTip();
            ttCheckBox.SetToolTip(checkBox1, "Include the slide background while performing Auto Zoom.");

            ToolTip ttCheckBox2 = new ToolTip();
            ttCheckBox2.SetToolTip(checkBox2, "Use separate slides for individual animation effects of Zoom to Area.");
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            SettingsHandler(this.checkBox1.Checked, this.checkBox2.Checked);
            this.Close();
        }
    }
}
