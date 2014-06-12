using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Views
{
    public partial class ColorPickerForm : Form
    {
        public ColorPickerForm()
        {
            InitializeComponent();
        }

        public ColorPickerForm(PowerPoint.ShapeRange selectedShapes)
            : this()
        {

        }

        private void ColorPickerForm_Load(object sender, EventArgs e)
        {

        }
    }
}
