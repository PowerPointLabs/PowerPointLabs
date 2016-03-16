using System;
using System.Windows.Forms;

namespace PowerPointLabs.Views
{
    public partial class Form1 : Form
    {
        public string newName;
        private Ribbon1 ribbon;
        public Form1()
        {
            InitializeComponent();
        }

        public Form1(Ribbon1 parentRibbon, String oldName)
            : this()
        {
            ribbon = parentRibbon;
            newName = oldName;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            this.newName = this.textBox1.Text;
            ribbon.ShapeNameEdited(this.newName);
            this.Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.textBox1.Text = newName;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
