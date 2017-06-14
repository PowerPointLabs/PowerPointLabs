using System;
using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Views
{
    public partial class EditNameDialogBox : Form
    {
        public Shape shape;
        private Ribbon1 ribbon;
        public EditNameDialogBox()
        {
            InitializeComponent();
        }

        public EditNameDialogBox(Ribbon1 parentRibbon, Shape selectedShape)
            : this()
        {
            ribbon = parentRibbon;
            shape = selectedShape;
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            this.shape.Name = this.editNameTextBox.Text;
            this.Close();
        }

        private void EditNameDialogBox_Load(object sender, EventArgs e)
        {
            this.editNameTextBox.Text = shape.Name;
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
