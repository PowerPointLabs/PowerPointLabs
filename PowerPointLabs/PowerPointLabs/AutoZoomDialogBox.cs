using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PowerPointLabs
{
    public partial class AutoZoomDialogBox : Form
    {
        public bool slideBackgroundChecked;
        private Ribbon1 ribbon;
        public AutoZoomDialogBox()
        {
            InitializeComponent();
        }

        public AutoZoomDialogBox(Ribbon1 parentRibbon, bool backgroundChecked)
            : this()
        {
            ribbon = parentRibbon;
            slideBackgroundChecked = backgroundChecked;
        }

        private void AutoZoomDialogBox_Load(object sender, EventArgs e)
        {
            this.checkBox1.Checked = slideBackgroundChecked;
            ToolTip ttCheckBox = new ToolTip();
            ttCheckBox.SetToolTip(checkBox1, "Include the slide background while performing Auto Zoom.");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.slideBackgroundChecked = this.checkBox1.Checked;
            ribbon.ZoomPropertiesEdited(slideBackgroundChecked);
            this.Dispose();
        }
    }
}
