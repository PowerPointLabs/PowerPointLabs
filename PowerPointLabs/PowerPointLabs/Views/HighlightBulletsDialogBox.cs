using System;
using System.Drawing;
using System.Windows.Forms;

namespace PowerPointLabs.Views
{
    public partial class HighlightBulletsDialogBox : Form
    {
        public delegate void UpdateSettingsDelegate(Color highlightColor, Color defaultColor, Color backgroundColor);
        public UpdateSettingsDelegate SettingsHandler;
        public HighlightBulletsDialogBox()
        {
            InitializeComponent();
        }

        public HighlightBulletsDialogBox(Color defaultHighlightColor, Color defaultTextColor, Color defaultBackgroundColor)
            : this()
        {
            this.pictureBox1.BackColor = defaultHighlightColor;
            this.pictureBox2.BackColor = defaultTextColor;
            this.pictureBox3.BackColor = defaultBackgroundColor;
        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            ColorDialog colorDialog = new ColorDialog();
            colorDialog.Color = this.pictureBox1.BackColor;
            colorDialog.FullOpen = true;
            if (colorDialog.ShowDialog() != DialogResult.Cancel)
            {
                this.pictureBox1.BackColor = colorDialog.Color;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SettingsHandler(this.pictureBox1.BackColor, this.pictureBox2.BackColor, this.pictureBox3.BackColor);
            this.Close();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            ColorDialog colorDialog = new ColorDialog();
            colorDialog.Color = this.pictureBox2.BackColor;
            colorDialog.FullOpen = true;
            if (colorDialog.ShowDialog() != DialogResult.Cancel)
            {
                this.pictureBox2.BackColor = colorDialog.Color;
            }
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            ColorDialog colorDialog = new ColorDialog();
            colorDialog.Color = this.pictureBox3.BackColor;
            colorDialog.FullOpen = true;
            if (colorDialog.ShowDialog() != DialogResult.Cancel)
            {
                this.pictureBox3.BackColor = colorDialog.Color;
            }
        }
    }
}
