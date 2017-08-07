using System.Windows;

using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ShortcutsLab.Views
{
    /// <summary>
    /// Interaction logic for EditNameDialogBox.xaml
    /// </summary>
    public partial class EditNameDialogBox
    {
        public Shape SelectedShape { get; private set; }

        public EditNameDialogBox(Shape refShape = null)
        {
            InitializeComponent();

            SelectedShape = refShape;
            textBoxNameInput.Text = SelectedShape.Name;
            textBoxNameInput.SelectAll();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            SelectedShape.Name = textBoxNameInput.Text;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
