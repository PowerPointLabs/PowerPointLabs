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

        private void button1_Click(object sender, EventArgs e)
        {
            this.newName = this.textBox1.Text;
            ribbon.nameEdited(this.newName);
            this.Dispose();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.textBox1.Text = newName;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }
    }
}
