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
        public bool singleSlideChecked;
        private Ribbon1 ribbon;
        public AutoZoomDialogBox()
        {
            InitializeComponent();
        }

        public AutoZoomDialogBox(Ribbon1 parentRibbon, bool backgroundChecked, bool singleChecked)
            : this()
        {
            ribbon = parentRibbon;
            slideBackgroundChecked = backgroundChecked;
            singleSlideChecked = singleChecked;
        }

        private void AutoZoomDialogBox_Load(object sender, EventArgs e)
        {
            this.checkBox1.Checked = slideBackgroundChecked;
            this.checkBox2.Checked = singleSlideChecked;

            ToolTip ttCheckBox = new ToolTip();
            ttCheckBox.SetToolTip(checkBox1, "Include the slide background while performing Auto Zoom.");

            ToolTip ttCheckBox2 = new ToolTip();
            ttCheckBox2.SetToolTip(checkBox2, "Perform all zoom animations within a single slide.\nThis may result in slower loading time for the animation slide.");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.slideBackgroundChecked = this.checkBox1.Checked;
            this.singleSlideChecked = this.checkBox2.Checked;
            ribbon.ZoomPropertiesEdited(slideBackgroundChecked, singleSlideChecked);
            this.Dispose();
        }
    }
}
