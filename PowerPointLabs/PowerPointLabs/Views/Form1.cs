using System;
using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Views
{
    public partial class Form1 : Form
    {
        public Shape shape;
        private Ribbon1 ribbon;
        public Form1()
        {
            InitializeComponent();
        }

        public Form1(Ribbon1 parentRibbon, Shape selectedShape)
            : this()
        {
            ribbon = parentRibbon;
            shape = selectedShape;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            this.shape.Name = this.textBox1.Text;
            this.Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.textBox1.Text = shape.Name;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
