using System;
using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Views
{
    public partial class EditNameDialogBox : Form
    {
        public Shape SelectedShape { get; private set; }

        public EditNameDialogBox()
        {
            InitializeComponent();
        }

        public EditNameDialogBox(Ribbon1 parentRibbon, Shape selectedShape)
            : this()
        {
            SelectedShape = selectedShape;
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            SelectedShape.Name = editNameTextBox.Text;
            this.Close();
        }

        private void EditNameDialogBox_Load(object sender, EventArgs e)
        {
            editNameTextBox.Text = SelectedShape.Name;
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
